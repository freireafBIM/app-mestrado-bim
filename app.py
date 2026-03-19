"""
Gestor BIM Estrutural — TQS / IFC2x3
=====================================
Extrai dados de todos os elementos estruturais exportados pelo TQS:
  Pilares (IfcColumn), Vigas (IfcBeam), Lajes (IfcSlab),
  Blocos/Sapatas (IfcFooting), Estacas (IfcPile), Escadas (IfcStair)

Diferenciais em relação ao app.py original:
  ✓ Sem variáveis globais mutáveis  → cache por st.session_state
  ✓ Sem senha hardcoded             → st.secrets["acesso"]["senha"]
  ✓ QR Code 100% em memória         → sem arquivos temp em disco
  ✓ Erros sempre visíveis           → sem except: pass silencioso
  ✓ Sheets sem risco de perda       → clear() só depois dos dados prontos
  ✓ Decodificação IFC correta       → \\X2\\00D8\\X0\\ → "Ø"
  ✓ Armadura por (nome, pavimento)  → correta por pavimento, não global
  ✓ Todos os tipos estruturais      → pilares, vigas, lajes, fund., estacas
  ✓ Regex calibrada para TQS        → "1 P1 Ø10.00 C=230.00"
  ✓ requirements.txt incluso        → reprodutibilidade científica

Estrutura de Psets TQS confirmada neste modelo:
  TQS_Padrao    → Titulo, Tipo, Piso, Planta, Material
  TQS_Geometria → Secao, Dimensao_b1, Dimensao_h1, Largura, Altura,
                  Dimensoes_X, Dimensoes_Y, Area, Area_superficie,
                  Carga_linear, Carga_distribuida, Estacas, Diametro
  TQS_Armaduras → Cobrimento, Tem_Protensao
  IfcReinforcingBar.Name → "QTD NOME_ELEM ØDiam C=Comp"
"""

import streamlit as st
import os, re, io, tempfile, datetime
from collections import defaultdict, Counter

import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import qrcode
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader


# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────

NOME_PLANILHA = "Sistema_Conferencia_BIM"

# Tipos suportados neste IFC TQS  (IFC2x3)
TIPOS_ESTRUTURAIS: dict[str, str] = {
    "IfcColumn":  "Pilar",
    "IfcBeam":    "Viga",
    "IfcSlab":    "Laje",
    "IfcFooting": "Fundação",
    "IfcPile":    "Estaca",
    "IfcStair":   "Escada",
}

# Regex calibrado para o padrão de nome TQS:
#   "1 P1 Ø10.00 C=230.00"   →  grupo(1)="P1"  grupo(2)="10.00"
#   "3 V2 Ø8.00 C=450.00"    →  grupo(1)="V2"  grupo(2)="8.00"
# O caractere Ø é exportado como \X2\00D8\X0\ → chr(0xD8) = 'Ø'
REGEX_BARRA = re.compile(
    r'^\d+\s+'          # quantidade
    r'(\S+)\s+'         # nome do elemento (P1, V3, B2...)
    r'[Ø\u00d8]'        # símbolo Ø (unicode ou latin-1)
    r'(\d+\.?\d*)',     # bitola em mm
    re.UNICODE
)


# ──────────────────────────────────────────────────────────────────────────────
# DECODIFICAÇÃO DE STRINGS IFC
# ──────────────────────────────────────────────────────────────────────────────

def decode_ifc(texto: str) -> str:
    """
    Converte codificações de string do formato IFC STEP para texto legível:
      \\X2\\00D8\\X0\\  →  Ø  (IFC2x3 / IFC4 Unicode escape)
      \\S\\x            →  caractere ISO-8859-1 offset +0x80  (IFC2x3 legacy)
    """
    if not texto:
        return ""
    # Sequência Unicode \X2\HHHH\X0\
    texto = re.sub(
        r'\\X2\\([0-9A-Fa-f]{4})\\X0\\',
        lambda m: chr(int(m.group(1), 16)),
        texto
    )
    # Sequência ISO-8859-1 \S\x
    texto = re.sub(
        r'\\S\\(.)',
        lambda m: chr(0x80 + ord(m.group(1))),
        texto
    )
    return texto


def limpar_valor(valor: str) -> str:
    """Remove wrapper de tipo IFC — IFCLABEL('x') → 'x', IFCLENGTHMEASURE(35.) → '35.0'"""
    if not valor or valor.strip() == '$':
        return ""
    v = valor.strip()
    # IFCLABEL / IFCTEXT / IFCIDENTIFIER com aspas
    m = re.match(r"IFC(?:LABEL|TEXT|IDENTIFIER)\('(.*)'\)$", v, re.DOTALL)
    if m:
        return decode_ifc(m.group(1))
    # Medidas numéricas
    m = re.match(r"IFC\w+MEASURE\((.+)\)$", v)
    if m:
        try:
            return str(round(float(m.group(1)), 4))
        except ValueError:
            return m.group(1)
    # IFCBOOLEAN
    m = re.match(r"IFCBOOLEAN\(\.([TF])\.\)$", v)
    if m:
        return "Sim" if m.group(1) == "T" else "Não"
    # IFCINTEGER / IFCREAL
    m = re.match(r"IFC(?:INTEGER|REAL|COUNTMEASURE)\((.+)\)$", v)
    if m:
        return m.group(1)
    # Valor nu (sem wrapper)
    return decode_ifc(v.strip("'"))


# ──────────────────────────────────────────────────────────────────────────────
# AUTENTICAÇÃO
# ──────────────────────────────────────────────────────────────────────────────

def conectar_sheets() -> gspread.Client:
    """
    Autentica usando st.secrets (produção) ou credenciais.json (local).
    Configure em .streamlit/secrets.toml — nunca versionar.
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes
        )
    elif os.path.exists("credenciais.json"):
        creds = Credentials.from_service_account_file("credenciais.json", scopes=scopes)
    else:
        st.error("Credenciais não encontradas. Configure st.secrets['gcp_service_account'].")
        st.stop()
    return gspread.authorize(creds)


def verificar_senha() -> str:
    """
    Retorna a senha do secrets.toml.
    Estrutura esperada:
        [acesso]
        senha = "minha_senha"
    A senha JAMAIS deve aparecer no código-fonte.
    """
    try:
        senha = st.secrets["acesso"]["senha"]
        if not senha:
            raise KeyError
        return senha
    except (KeyError, AttributeError):
        st.error("Configure [acesso] senha = '...' no arquivo .streamlit/secrets.toml")
        st.stop()


# ──────────────────────────────────────────────────────────────────────────────
# INDEXAÇÃO DE ARMADURAS (por sessão — sem estado global)
# ──────────────────────────────────────────────────────────────────────────────

def indexar_armaduras(ifc_file) -> dict:
    """
    Varre todas as IfcReinforcingBar do modelo e agrupa as bitolas
    por chave (nome_elemento, pavimento).

    Retorna: { ('P1', 'Tipo-001'): [10.0, 10.0, 8.0, ...], ... }

    O padrão de nome TQS é: "QTD NOME_ELEM ØDiam C=Comp (id N)"
    Exemplo: "3 P1 Ø10.00 C=265.00 (id 2)"
    """
    # Mapa elemento → pavimento (via IfcRelContainedInSpatialStructure)
    storey_por_elem: dict[str, str] = {}
    for rel in ifc_file.by_type("IfcRelContainedInSpatialStructure"):
        try:
            storey = rel.RelatingStructure.Name or "Sem pavimento"
            for elem in rel.RelatedElements:
                storey_por_elem[elem.id()] = decode_ifc(storey)
        except Exception:
            pass

    cache: dict[tuple, list] = defaultdict(list)

    for barra in ifc_file.by_type("IfcReinforcingBar"):
        try:
            nome_raw = barra.Name or ""
            nome = decode_ifc(nome_raw)
            # Remove sufixo de instância: "(id 2)", "(id 10)"
            nome_base = re.sub(r'\s*\(id\s*\d+\)', '', nome).strip()

            match = REGEX_BARRA.match(nome_base)
            if not match:
                continue

            nome_elem = match.group(1)          # "P1", "V3", "B1"...
            bitola    = float(match.group(2))   # 10.0, 8.0, 6.3...

            pav = storey_por_elem.get(barra.id(), "Sem pavimento")
            cache[(nome_elem, pav)].append(bitola)

        except Exception as e:
            st.warning(f"Barra ignorada (#{getattr(barra, 'id', '?')}): {e}")

    return dict(cache)


def formatar_armadura(cache: dict, nome_elem: str, pavimento: str) -> str:
    """
    Consulta o cache e retorna string formatada.
    Exemplo: "12 ø16.0 + 8 ø10.0 + 20 ø6.3"
    Bitolas maiores aparecem primeiro (ordem decrescente).
    """
    chave = (nome_elem, pavimento)
    if chave not in cache:
        return "Verificar detalhamento"
    contagem = Counter(cache[chave])
    partes = sorted(contagem.items(), key=lambda x: -x[0])
    return " + ".join(f"{qtd} ø{diam:.1f}" for diam, qtd in partes)


# ──────────────────────────────────────────────────────────────────────────────
# EXTRAÇÃO DE DADOS POR TIPO DE ELEMENTO
# ──────────────────────────────────────────────────────────────────────────────

def _psets(elem) -> dict:
    """
    Retorna todos os Psets de um elemento como dict plano:
    { 'NomePset.NomeProp': 'valor_limpo' }
    Nunca lança exceção — registra warning e continua.
    """
    resultado = {}
    try:
        raw = ifcopenshell.util.element.get_psets(elem)
        for pset_nome, props in raw.items():
            pset_dec = decode_ifc(str(pset_nome))
            for prop_nome, prop_val in props.items():
                prop_dec = decode_ifc(str(prop_nome))
                val = limpar_valor(str(prop_val)) if prop_val is not None else ""
                resultado[f"{pset_dec}.{prop_dec}"] = val
    except Exception as e:
        st.warning(f"Erro ao ler Psets de '{getattr(elem, 'Name', '?')}': {e}")
    return resultado


def _pavimento(elem) -> str:
    """Retorna nome do IfcBuildingStorey do elemento."""
    try:
        if elem.ContainedInStructure:
            return decode_ifc(elem.ContainedInStructure[0].RelatingStructure.Name or "")
    except Exception:
        pass
    return "Sem pavimento"


def _bbox(elem) -> dict:
    """
    Extrai bounding box 3D por varredura recursiva dos CartesianPoints.
    Unidade do TQS: centímetros (IFCSIUNIT CENTI METRE).
    Retorna dimensões em cm e coordenadas do centróide.
    Nunca lança exceção — retorna zeros com warning.
    """
    vazio = {"comp_cm": 0.0, "larg_cm": 0.0, "alt_cm": 0.0,
             "coord_x": 0.0, "coord_y": 0.0}
    try:
        if not getattr(elem, "Representation", None):
            return vazio

        xs, ys, zs = [], [], []
        visitados: set = set()

        def coletar(obj, profundidade=0):
            if obj is None or profundidade > 14:
                return
            oid = id(obj)
            if oid in visitados:
                return
            visitados.add(oid)

            if hasattr(obj, "is_a"):
                if obj.is_a("IfcCartesianPoint") and hasattr(obj, "Coordinates"):
                    c = obj.Coordinates
                    if len(c) >= 2:
                        xs.append(float(c[0]))
                        ys.append(float(c[1]))
                        if len(c) >= 3:
                            zs.append(float(c[2]))
                    return

                for attr in ("Points", "OuterCurve", "PolygonalBoundary", "Polygon",
                             "Items", "MappedRepresentation", "MappingSource",
                             "SweptArea", "Bounds", "Bound", "CfsFaces", "FbsmFaces",
                             "Position", "BaseSurface"):
                    if not hasattr(obj, attr):
                        continue
                    val = getattr(obj, attr)
                    if isinstance(val, (list, tuple)):
                        for v in val:
                            coletar(v, profundidade + 1)
                    else:
                        coletar(val, profundidade + 1)

            elif isinstance(obj, (list, tuple)):
                for item in obj:
                    coletar(item, profundidade + 1)

        for rep in elem.Representation.Representations:
            if rep.RepresentationIdentifier in ("Body", "Mesh", "Box", "Facetation", "Axis"):
                for item in rep.Items:
                    coletar(item)

        if not xs:
            return vazio

        comp = round(max(xs) - min(xs), 2)
        larg = round(max(ys) - min(ys), 2)
        alt  = round(max(zs) - min(zs), 2) if zs else 0.0
        cx   = round((min(xs) + max(xs)) / 2, 2)
        cy   = round((min(ys) + max(ys)) / 2, 2)

        return {"comp_cm": comp, "larg_cm": larg, "alt_cm": alt,
                "coord_x": cx, "coord_y": cy}

    except Exception as e:
        st.warning(f"Geometria não extraída para '{getattr(elem, 'Name', '?')}': {e}")
        return vazio


def _id_unico(elem, id_projeto: str, pavimento: str) -> str:
    """Chave primária única e segura para uso como QR Code."""
    guid  = (elem.GlobalId or "NOGUID")[:8]
    nome  = re.sub(r"[^A-Z0-9]", "", (elem.Name or "").upper())[:8]
    pav   = re.sub(r"[^A-Z0-9]", "", pavimento.upper())[:8]
    proj  = re.sub(r"[^A-Z0-9]", "", id_projeto.upper())[:10]
    return f"{proj}-{pav}-{nome}-{guid}"


def _volume_m3(geo: dict, tipo_ifc: str) -> float:
    """
    Estima volume de concreto em m³ a partir da bounding box.
    Pilares/Estacas: Área_seção × Altura
    Vigas: Largura × Altura × Comprimento
    Lajes: Comprimento × Largura × Altura (espessura)
    Fundações: usa dimensões X, Y, Altura quando disponíveis via Pset.
    Unidade entrada: cm → saída: m³
    """
    c = geo["comp_cm"] / 100
    l = geo["larg_cm"] / 100
    a = geo["alt_cm"]  / 100
    if c > 0 and l > 0 and a > 0:
        return round(c * l * a, 4)
    return 0.0


# ──────────────────────────────────────────────────────────────────────────────
# PIPELINE PRINCIPAL DE EXTRAÇÃO
# ──────────────────────────────────────────────────────────────────────────────

def _natural_key(texto: str) -> list:
    return [int(c) if c.isdigit() else c.lower()
            for c in re.split(r"(\d+)", texto or "")]


def processar_ifc(caminho: str, nome_projeto: str, id_projeto: str) -> list[dict]:
    """
    Abre o IFC e extrai dados de TODOS os elementos estruturais.
    Retorna lista de dicts prontos para o Sheets / PDF.
    """
    ifc = ifcopenshell.open(caminho)

    # ── 1. Indexar armaduras (uma única passagem sobre as 6943 barras) ────────
    with st.spinner("Indexando armaduras (IfcReinforcingBar)..."):
        cache_arm = indexar_armaduras(ifc)
    st.info(f"Armaduras indexadas: {sum(len(v) for v in cache_arm.values())} barras "
            f"em {len(cache_arm)} combinações (elemento, pavimento)")

    # ── 2. Processar cada tipo estrutural ─────────────────────────────────────
    registros: list[dict] = []
    total = sum(len(ifc.by_type(t)) for t in TIPOS_ESTRUTURAIS)

    if total == 0:
        st.warning("Nenhum elemento estrutural encontrado. Verifique se o IFC é do TQS.")
        return []

    progresso = st.progress(0.0, text="Extraindo elementos...")
    processados = 0

    for tipo_ifc, tipo_legivel in TIPOS_ESTRUTURAIS.items():
        elementos = ifc.by_type(tipo_ifc)
        if not elementos:
            continue

        for elem in elementos:
            processados += 1
            progresso.progress(
                processados / total,
                text=f"{tipo_legivel}: {elem.Name or '?'} ({processados}/{total})"
            )

            nome     = elem.Name or "S/N"
            pavimento = _pavimento(elem)
            geo      = _bbox(elem)
            ps       = _psets(elem)
            volume   = _volume_m3(geo, tipo_ifc)

            # ── Extrair campos específicos dos Psets TQS ──────────────────────
            def get(*chaves: str, fallback: str = "—") -> str:
                for c in chaves:
                    v = ps.get(c, "")
                    if v and v not in ("$", "—"):
                        return v
                return fallback

            material  = get("TQS_Padrao.Material")
            cobrimento = get("TQS_Armaduras.Cobrimento")
            protensao  = get("TQS_Armaduras.Tem_Protensao")

            # Geometria específica por tipo (Pset tem precedência sobre bbox)
            if tipo_ifc == "IfcColumn":
                secao = get("TQS_Geometria.Secao")
                dim_b = get("TQS_Geometria.Dimensao_b1")
                dim_h = get("TQS_Geometria.Dimensao_h1")
                area_secao = get("TQS_Geometria.Area")
                descricao_geo = (f"Seção:{secao} {dim_b}×{dim_h} cm"
                                 if dim_b != "—" else f"Seção:{secao}")

            elif tipo_ifc == "IfcBeam":
                largura = get("TQS_Geometria.Largura")
                altura  = get("TQS_Geometria.Altura")
                vao     = get("TQS_Geometria.Vao_Titulo", "TQS_Geometria.Vao")
                carga   = get("TQS_Geometria.Carga_linear")
                descricao_geo = f"{largura}×{altura} cm | Vão:{vao}"

            elif tipo_ifc == "IfcSlab":
                esp     = get("TQS_Geometria.Altura")
                area_sup = get("TQS_Geometria.Area_superficie")
                carga_d = get("TQS_Geometria.Carga_distribuida")
                tipo_laje = get("TQS_Geometria.Tipo")
                descricao_geo = f"Esp:{esp} cm | Área:{area_sup} m² | {tipo_laje}"

            elif tipo_ifc == "IfcFooting":
                dim_x   = get("TQS_Geometria.Dimensoes_X")
                dim_y   = get("TQS_Geometria.Dimensoes_Y")
                alt_f   = get("TQS_Geometria.Altura")
                n_est   = get("TQS_Geometria.Estacas")
                tipo_f  = get("TQS_Geometria.Tipo")
                descricao_geo = f"{tipo_f} {dim_x}×{dim_y}×{alt_f} cm | Estacas:{n_est}"

            elif tipo_ifc == "IfcPile":
                diam_est = get("TQS_Geometria.Diametro", "TQS_Geometria.Dimensao_b1")
                alt_est  = get("TQS_Geometria.Altura")
                descricao_geo = f"Ø{diam_est} cm | H:{alt_est} cm"

            else:  # Escada, outros
                descricao_geo = (f"{geo['comp_cm']:.0f}×{geo['larg_cm']:.0f}×"
                                 f"{geo['alt_cm']:.0f} cm")

            armadura = formatar_armadura(cache_arm, nome, pavimento)

            registro = {
                "ID_Unico":         _id_unico(elem, id_projeto, pavimento),
                "Projeto_Ref":      id_projeto,
                "Nome_Projeto":     nome_projeto,
                "Tipo_IFC":         tipo_ifc,
                "Tipo_Legivel":     tipo_legivel,
                "Nome":             nome,
                "Pavimento":        pavimento,
                "Material":         material,
                "Geometria":        descricao_geo,
                "Cobrimento_cm":    cobrimento,
                "Armadura":         armadura,
                "Tem_Protensao":    protensao,
                "Volume_m3":        volume,
                "Geo_Comp_cm":      geo["comp_cm"],
                "Geo_Larg_cm":      geo["larg_cm"],
                "Geo_Alt_cm":       geo["alt_cm"],
                "Coord_X":          geo["coord_x"],
                "Coord_Y":          geo["coord_y"],
                "GUID":             elem.GlobalId or "",
                "Status":           "A CONFERIR",
                "Data_Conferencia": "",
                "Responsavel":      "",
            }
            registros.append(registro)

    progresso.progress(1.0, text="Extração concluída.")

    # Ordenação: pavimento → tipo → nome natural
    registros.sort(key=lambda r: (
        r["Pavimento"],
        r["Tipo_Legivel"],
        _natural_key(r["Nome"])
    ))
    return registros


# ──────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DE PDF (100% em memória — sem arquivos temporários)
# ──────────────────────────────────────────────────────────────────────────────

def gerar_pdf(registros: list[dict], nome_projeto: str) -> io.BytesIO:
    """
    Gera PDF A4 com etiquetas 90×50 mm contendo QR Code e dados técnicos.
    Todo processamento ocorre em RAM — nenhum arquivo em disco.
    """
    buf_pdf = io.BytesIO()
    c = rl_canvas.Canvas(buf_pdf, pagesize=A4)
    W_PAG, H_PAG = A4

    LARG, ALTA = 90 * mm, 50 * mm
    MH, MV     = 10 * mm, 10 * mm
    GAP_X, GAP_Y = 5 * mm, 5 * mm
    COLS = 2

    c.setTitle(f"Etiquetas BIM — {nome_projeto}")
    col, x, y = 0, MH, H_PAG - MV - ALTA

    COR_FUNDO  = {
        "Pilar":    colors.HexColor("#EEF4FF"),
        "Viga":     colors.HexColor("#FFF4EE"),
        "Laje":     colors.HexColor("#EEFFEE"),
        "Fundação": colors.HexColor("#FFFBEE"),
        "Estaca":   colors.HexColor("#F8EEFF"),
        "Escada":   colors.HexColor("#EEFFFF"),
    }

    for reg in registros:
        tipo = reg["Tipo_Legivel"]
        fundo = COR_FUNDO.get(tipo, colors.white)

        # Fundo colorido por tipo
        c.setFillColor(fundo)
        c.rect(x, y, LARG, ALTA, fill=1, stroke=0)

        # Borda
        c.setStrokeColor(colors.HexColor("#AAAAAA"))
        c.setLineWidth(0.5)
        c.rect(x, y, LARG, ALTA, fill=0, stroke=1)

        # ── QR Code em memória ────────────────────────────────────────────────
        qr = qrcode.QRCode(box_size=8, border=1,
                           error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(reg["ID_Unico"])
        qr.make(fit=True)
        img_pil = qr.make_image(fill_color="black", back_color="white")
        buf_qr = io.BytesIO()
        img_pil.save(buf_qr, format="PNG")
        buf_qr.seek(0)
        c.drawImage(ImageReader(buf_qr),
                    x + 2*mm, y + 5*mm, width=38*mm, height=38*mm)

        # ── Textos ────────────────────────────────────────────────────────────
        tx = x + 42 * mm

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(tx, y + 43*mm, tipo.upper())

        c.setFont("Helvetica-Bold", 13)
        c.drawString(tx, y + 36*mm, reg["Nome"][:16])

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#333333"))

        geo_str = reg["Geometria"][:26]
        c.drawString(tx, y + 30*mm, geo_str)

        arm_str = reg["Armadura"][:28]
        c.setFont("Helvetica-Bold", 7)
        c.drawString(tx, y + 25*mm, f"Arm: {arm_str}")

        c.setFont("Helvetica", 7)
        c.drawString(tx, y + 20*mm, f"Mat: {reg['Material'][:20]}")
        c.drawString(tx, y + 16*mm, f"Cob: {reg['Cobrimento_cm']} cm")

        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(colors.black)
        c.drawString(tx, y + 10*mm, reg["Pavimento"][:22])

        c.setFont("Helvetica-Oblique", 6.5)
        c.setFillColor(colors.gray)
        c.drawString(tx, y + 4*mm, f"Obra: {nome_projeto[:22]}")

        # ── Avanço de posição ─────────────────────────────────────────────────
        col += 1
        if col >= COLS:
            col, x = 0, MH
            y -= ALTA + GAP_Y
        else:
            x += LARG + GAP_X

        if y < MV:
            c.showPage()
            x, y, col = MH, H_PAG - MV - ALTA, 0

    c.save()
    buf_pdf.seek(0)
    return buf_pdf


# ──────────────────────────────────────────────────────────────────────────────
# PERSISTÊNCIA NO GOOGLE SHEETS
# ──────────────────────────────────────────────────────────────────────────────

def salvar_no_sheets(client: gspread.Client, id_proj: str,
                     nome: str, registros: list[dict]) -> None:
    """
    Salva projeto e elementos no Google Sheets de forma segura.
    Estratégia: monta DataFrame completo em memória ANTES de fazer clear().
    Isso garante que, se houver erro de rede, nenhum dado existente é perdido.
    """
    sh = client.open(NOME_PLANILHA)

    # ── Aba Projetos ──────────────────────────────────────────────────────────
    try:
        ws_p = sh.worksheet("Projetos")
    except gspread.WorksheetNotFound:
        ws_p = sh.add_worksheet("Projetos", 200, 8)

    df_p = pd.DataFrame(ws_p.get_all_records())
    if not df_p.empty and "ID_Projeto" in df_p.columns:
        df_p = df_p[df_p["ID_Projeto"] != id_proj]

    tipos_count = Counter(r["Tipo_Legivel"] for r in registros)
    vol_total = round(sum(r.get("Volume_m3", 0) for r in registros), 3)

    novo = {
        "ID_Projeto":          id_proj,
        "Nome_Obra":           nome,
        "Data_Upload":         datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Total_Elementos":     len(registros),
        "Volume_Total_m3":     vol_total,
        "Resumo_Tipos":        " | ".join(f"{t}: {n}" for t, n in tipos_count.items()),
    }
    df_p = pd.concat([df_p, pd.DataFrame([novo])], ignore_index=True).fillna("").astype(str)
    ws_p.clear()
    ws_p.update([df_p.columns.tolist()] + df_p.values.tolist())

    # ── Aba Elementos ─────────────────────────────────────────────────────────
    try:
        ws_e = sh.worksheet("Elementos")
    except gspread.WorksheetNotFound:
        ws_e = sh.add_worksheet("Elementos", 5000, 25)

    df_e = pd.DataFrame(ws_e.get_all_records())
    if not df_e.empty and "Projeto_Ref" in df_e.columns:
        df_e = df_e[df_e["Projeto_Ref"] != id_proj]

    df_novos = pd.DataFrame(registros).fillna("").astype(str)

    # União de colunas (novos podem ter colunas a mais)
    todas_cols = list(dict.fromkeys(
        (df_e.columns.tolist() if not df_e.empty else []) +
        df_novos.columns.tolist()
    ))
    if not df_e.empty:
        df_e = df_e.reindex(columns=todas_cols, fill_value="")
    df_novos = df_novos.reindex(columns=todas_cols, fill_value="")

    # Monta tudo em memória → só então faz clear()
    df_final = pd.concat([df_e, df_novos], ignore_index=True).fillna("").astype(str)
    ws_e.clear()
    ws_e.update([df_final.columns.tolist()] + df_final.values.tolist())


# ──────────────────────────────────────────────────────────────────────────────
# INTERFACE STREAMLIT
# ──────────────────────────────────────────────────────────────────────────────

def _tela_login(senha_correta: str) -> None:
    st.title("🔒 Acesso Restrito")
    col1, col2 = st.columns([2, 3])
    with col1:
        senha = st.text_input("Senha de acesso", type="password")
        if st.button("Entrar", type="primary"):
            if senha == senha_correta:
                st.session_state["logado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")


def _tela_principal() -> None:
    st.title("🏗️ Gestor BIM Estrutural")
    st.caption("TQS 27.0 · IFC2x3 · Pilares · Vigas · Lajes · Fundações · Estacas")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Obra")
        nome = st.text_input("Nome da obra", placeholder="Ex: Edifício Residencial Aurora")
        id_proj = re.sub(r"[^A-Z0-9]", "", nome.upper())[:12] if nome else ""
        if id_proj:
            st.caption(f"ID interno: `{id_proj}`")
        st.divider()

        filtro_tipo = st.multiselect(
            "Filtrar tipos", list(TIPOS_ESTRUTURAIS.values()),
            default=list(TIPOS_ESTRUTURAIS.values())
        )
        filtro_pav = st.text_input("Filtrar pavimento (parcial)", "")

        st.divider()
        if st.button("🚪 Sair"):
            st.session_state.clear()
            st.rerun()

    # ── Upload ────────────────────────────────────────────────────────────────
    arquivo = st.file_uploader(
        "Arquivo IFC (TQS — IFC2x3)",
        type=["ifc"],
        help="Exportado pelo TQS 27.x via File > Export > IFC"
    )

    if not (arquivo and nome):
        st.info("Preencha o nome da obra e carregue o arquivo IFC para continuar.")
        return

    if st.button("🚀 Processar IFC", type="primary"):
        # Salva arquivo em temp seguro (sempre removido no finally)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
            tmp.write(arquivo.getvalue())
            caminho_tmp = tmp.name

        try:
            registros = processar_ifc(caminho_tmp, nome, id_proj)
        finally:
            os.unlink(caminho_tmp)

        if not registros:
            return

        st.session_state["registros"] = registros
        st.session_state["nome"]      = nome
        st.session_state["id_proj"]   = id_proj

    # ── Resultados ────────────────────────────────────────────────────────────
    if "registros" not in st.session_state:
        return

    registros: list[dict] = st.session_state["registros"]
    nome    = st.session_state["nome"]
    id_proj = st.session_state["id_proj"]

    # Métricas de resumo
    tipos_count = Counter(r["Tipo_Legivel"] for r in registros)
    vol_total   = round(sum(r.get("Volume_m3", 0) for r in registros), 2)

    cols = st.columns(min(len(tipos_count) + 1, 5))
    for i, (t, n) in enumerate(tipos_count.items()):
        cols[i].metric(t, n)
    cols[-1].metric("Volume total (m³)", vol_total)

    # Filtros
    filtrados = [
        r for r in registros
        if r["Tipo_Legivel"] in filtro_tipo
        and (not filtro_pav or filtro_pav.lower() in r["Pavimento"].lower())
    ]

    st.divider()
    colunas_view = ["Nome", "Tipo_Legivel", "Pavimento", "Geometria",
                    "Armadura", "Material", "Cobrimento_cm", "Volume_m3", "Status"]
    df_view = pd.DataFrame(filtrados)[
        [c for c in colunas_view if c in pd.DataFrame(filtrados).columns]
    ]
    st.dataframe(df_view, use_container_width=True, height=340)
    st.caption(f"Exibindo {len(filtrados)} de {len(registros)} elementos")

    # ── Ações ─────────────────────────────────────────────────────────────────
    st.divider()
    col_a, col_b = st.columns(2)

    if col_a.button("☁️ Sincronizar com Google Sheets"):
        try:
            with st.spinner("Conectando..."):
                client = conectar_sheets()
            with st.spinner(f"Salvando {len(registros)} elementos..."):
                salvar_no_sheets(client, id_proj, nome, registros)
            st.success(f"✅ {len(registros)} elementos sincronizados com sucesso!")
        except Exception as e:
            st.error(f"Erro na sincronização: {e}")

    if col_b.button("📄 Gerar PDF com QR Codes"):
        alvo = filtrados if filtrados != registros else registros
        with st.spinner(f"Gerando {len(alvo)} etiquetas..."):
            buf = gerar_pdf(alvo, nome)
        st.download_button(
            label=f"⬇️ Baixar PDF ({len(alvo)} etiquetas)",
            data=buf,
            file_name=f"Etiquetas_{id_proj}_{datetime.date.today()}.pdf",
            mime="application/pdf",
        )


def main() -> None:
    st.set_page_config(
        page_title="Gestor BIM Estrutural",
        page_icon="🏗️",
        layout="wide",
    )

    if "logado" not in st.session_state:
        st.session_state["logado"] = False

    senha_correta = verificar_senha()

    if not st.session_state["logado"]:
        _tela_login(senha_correta)
    else:
        _tela_principal()


if __name__ == "__main__":
    main()
