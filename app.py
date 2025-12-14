import streamlit as st
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
import io
import datetime
import re
from collections import Counter

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- CACHE GLOBAL ---
CACHE_ARMADURAS_POR_NOME = {}

# --- CONEXÃO ---
def conectar_google_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists(ARQUIVO_CREDENCIAIS):
        creds = Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
    else:
        st.error("ERRO CRÍTICO: Credenciais não encontradas.")
        st.stop()
    return gspread.authorize(creds)

def limpar_string(texto):
    if not texto: return "X"
    return "".join(e for e in str(texto) if e.isalnum()).upper()

# --- ARMADURA (MÉTODO TQS / REGEX) ---
def indexar_todas_armaduras(ifc_file):
    global CACHE_ARMADURAS_POR_NOME
    CACHE_ARMADURAS_POR_NOME = {}
    barras = ifc_file.by_type('IfcReinforcingBar')
    
    for bar in barras:
        nome_completo = bar.Name 
        if not nome_completo: continue
        
        match_nome = re.search(r'^\d+\s+(P\d+)\s', nome_completo)
        
        if match_nome:
            nome_pilar = match_nome.group(1)
            bitola = 0.0
            
            try:
                resto_string = nome_completo.split(nome_pilar, 1)[1]
                match_bitola = re.search(r'([0-9]+\.[0-9]+)', resto_string)
                if match_bitola:
                    bitola = float(match_bitola.group(1))
            except:
                pass
            
            if bitola == 0.0 and hasattr(bar, "NominalDiameter") and bar.NominalDiameter:
                bitola = bar.NominalDiameter * 1000 
            
            qtd_barra = 1
            match_qtd = re.search(r'^(\d+)\s+P', nome_completo)
            if match_qtd:
                qtd_barra = int(match_qtd.group(1))

            if bitola > 0:
                if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
                    CACHE_ARMADURAS_POR_NOME[nome_pilar] = []
                for _ in range(qtd_barra):
                    CACHE_ARMADURAS_POR_NOME[nome_pilar].append(bitola)

def obter_armadura_do_cache(nome_pilar):
    if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
        return "Verificar Detalhamento"
    lista_bitolas = CACHE_ARMADURAS_POR_NOME[nome_pilar]
    c = Counter(lista_bitolas)
    return " + ".join([f"{qtd} ø{diam:.1f}" for diam, qtd in sorted(c.items(), key=lambda item: item[0], reverse=True)])

# --- GEOMETRIA E QUANTITATIVOS (MELHORIA BASEADA NO ARTIGO) ---

def extrair_dados_geometricos(pilar):
    """
    Retorna um dicionário com: Seção, Altura Estimada, Coordenadas (X,Y)
    """
    resultado = {
        "secao": "N/A", 
        "altura": 0.0, 
        "coord_x": 0.0, 
        "coord_y": 0.0,
        "largura_cm": 0.0,
        "profundidade_cm": 0.0
    }

    # 1. Tenta Psets TQS
    psets = ifcopenshell.util.element.get_psets(pilar)
    if 'TQS_Geometria' in psets:
        d = psets['TQS_Geometria']
        b = d.get('Dimensao_b1') or d.get('B')
        h = d.get('Dimensao_h1') or d.get('H')
        if b and h:
            vals = sorted([float(b), float(h)])
            if vals[0] < 3.0: vals = [v*100 for v in vals]
            resultado["secao"] = f"{vals[0]:.0f}x{vals[1]:.0f}"
            resultado["largura_cm"] = vals[0]
            resultado["profundidade_cm"] = vals[1]

    # 2. Varredura 3D (Bounding Box)
    if not pilar.Representation: return resultado
    
    pontos_x, pontos_y, pontos_z = [], [], []
    
    def coletar_pontos(item):
        if item is None: return
        if isinstance(item, (list, tuple)):
            for i in item: coletar_pontos(i)
            return
        if not hasattr(item, 'is_a'): return

        if item.is_a('IfcCartesianPoint') and hasattr(item, 'Coordinates'):
            c = item.Coordinates
            if len(c) >= 3:
                pontos_x.append(c[0])
                pontos_y.append(c[1])
                pontos_z.append(c[2])
            return

        atributos = ['Points', 'OuterCurve', 'PolygonalBoundary', 'FbsmFaces', 'CfsFaces', 'Bounds', 'Bound', 'Items', 'MappingSource', 'MappedRepresentation', 'Polygon', 'SweptArea']
        for attr in atributos:
            if hasattr(item, attr):
                coletar_pontos(getattr(item, attr))

    for rep in pilar.Representation.Representations:
        if rep.RepresentationIdentifier in ['Body', 'Mesh', 'Box', 'Facetation']:
            for item in rep.Items:
                coletar_pontos(item)
    
    if pontos_x and pontos_y and pontos_z:
        try:
            min_x, max_x = min(pontos_x), max(pontos_x)
            min_y, max_y = min(pontos_y), max(pontos_y)
            min_z, max_z = min(pontos_z), max(pontos_z)

            largura = max_x - min_x
            profundidade = max_y - min_y
            altura = max_z - min_z
            
            # Centroide aproximado
            resultado["coord_x"] = round((min_x + max_x) / 2, 2)
            resultado["coord_y"] = round((min_y + max_y) / 2, 2)
            resultado["altura"] = round(altura, 2) # Altura em metros (IFC padrão)

            # Ajuste de escala para seção (metros -> cm)
            if largura < 3.0: largura *= 100
            if profundidade < 3.0: profundidade *= 100
            
            dims = sorted([largura, profundidade])
            
            # Só atualiza se não achou via Pset
            if resultado["secao"] == "N/A":
                resultado["secao"] = f"{dims[0]:.0f}x{dims[1]:.0f}"
                resultado["largura_cm"] = dims[0]
                resultado["profundidade_cm"] = dims[1]
                
        except:
            pass
            
    return resultado

def extrair_material(pilar):
    """Tenta identificar o material (Concreto) associado"""
    try:
        material = ifcopenshell.util.element.get_material(pilar)
        if material:
            return material.Name
    except:
        pass
    return "Concreto" # Padrão se não achar

# --- ORDENAÇÃO ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    indexar_todas_armaduras(ifc_file)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name
        
        sufixo_pav = limpar_string(pavimento)
        sufixo_nome = limpar_string(nome)
        id_unico_pilar = f"{sufixo_nome}-{guid}-{sufixo_pav}-{id_projeto_input}"

        # Extração Geométrica Avançada (Com Volume)
        geo = extrair_dados_geometricos(pilar)
        armadura = obter_armadura_do_cache(nome)
        material = extrair_material(pilar)
        
        # Cálculo de Volume Estimado (Seção em cm * Altura em m)
        volume_estimado = 0.0
        if geo["largura_cm"] > 0 and geo["profundidade_cm"] > 0 and geo["altura"] > 0:
            area_m2 = (geo["largura_cm"] / 100) * (geo["profundidade_cm"] / 100)
            volume_estimado = round(area_m2 * geo["altura"], 3) # m³

        dados.append({
            'ID_Unico': id_unico_pilar,   
            'Projeto_Ref': id_projeto_input, 
            'Nome': nome, 
            'Secao': geo["secao"],
            'Altura_m': geo["altura"],
            'Volume_Concreto_m3': volume_estimado,
            'Material': material,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Coord_X': geo["coord_x"],
            'Coord_Y': geo["coord_y"],
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    dados.sort(key=lambda x: (x['Pavimento'], natural_keys(x['Nome'])))
    return dados

# --- PDF (LAYOUT RÍGIDO) ---
def gerar_pdf_memoria(dados_pilares, nome_projeto_legivel):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largura_pag, altura_pag = A4
    LARGURA_ETQ, ALTURA_ETQ = 90 * mm, 50 * mm
    MARGEM_ESQ, MARGEM_SUP = 10 * mm, 10 * mm
    ESPACO_X, ESPACO_Y = 5 * mm, 5 * mm
    
    x = MARGEM_ESQ
    y = altura_pag - MARGEM_SUP - ALTURA_ETQ
    coluna_atual = 0
    c.setTitle(f"Etiquetas - {nome_projeto_legivel}")
    
    for i, pilar in enumerate(dados_pilares):
        c.setLineWidth(1)
        c.setStrokeColor(colors.black)
        c.rect(x, y, LARGURA_ETQ, ALTURA_ETQ)
        
        qr = qrcode.QRCode(box_size=10, border=0)
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        temp_qr_path = f"temp_{i}.png"
        img_qr.save(temp_qr_path)
        c.drawImage(temp_qr_path, x + 3*mm, y + 7.5*mm, width=35*mm, height=35*mm)
        os.remove(temp_qr_path)
        
        texto_x = x + 42*mm
        c.setFont("Helvetica-Bold", 16)
        c.drawString(texto_x, y + 38*mm, f"PILAR: {pilar['Nome']}")
        
        # Dados Técnicos Expandidos na Etiqueta
        c.setFont("Helvetica", 10)
        c.drawString(texto_x, y + 31*mm, f"Sec: {pilar['Secao']} | H: {pilar['Altura_m']}m")
        c.drawString(texto_x, y + 27*mm, f"Vol: {pilar['Volume_Concreto_m3']} m³")
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(texto_x, y + 19*mm, f"{pilar['Pavimento']}")
        
        c.setFont("Helvetica-Oblique", 8)
        c.setFillColor(colors.gray)
        c.drawString(texto_x, y + 8*mm, f"Obra: {nome_projeto_legivel[:18]}...")
        c.setFillColor(colors.black)
        
        c.setDash(3, 3)
        c.setLineWidth(0.2)
        c.rect(x-1*mm, y-1*mm, LARGURA_ETQ+2*mm, ALTURA_ETQ+2*mm)
        c.setDash()
        
        coluna_atual += 1
        if coluna_atual > 1:
            coluna_atual = 0
            x = MARGEM_ESQ
            y -= (ALTURA_ETQ + ESPACO_Y)
        else:
            x += (LARGURA_ETQ + ESPACO_X)
        if y < MARGEM_SUP:
            c.showPage()
            y = altura_pag - MARGEM_SUP - ALTURA_ETQ
            x = MARGEM_ESQ
            coluna_atual = 0
    c.save()
    buffer.seek(0)
    return buffer

# --- FRONTEND ---
def main():
    st.set_page_config(page_title="Gestor BIM", page_icon="🏗️")
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    
    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        if st.text_input("Senha", type="password") == "bim123" and st.button("Entrar"):
            st.session_state['logado'] = True
            st.rerun()
        return

    st.title("🏗️ Gestor BIM 3.0 (TQS + Quantitativos)")
    nome = st.text_input("Nome da Obra", placeholder="Ex: Edifício Diogenes")
    id_proj = limpar_string(nome)
    f = st.file_uploader("Carregar IFC (TQS)", type=["ifc"])
    
    if f and nome:
        if st.button("🚀 PROCESSAR DADOS"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as t:
                    t.write(f.getvalue())
                    path = t.name
                
                with st.spinner('Minerando dados (Geometria + Armadura + Volume)...'):
                    dados = processar_ifc(path, id_proj)
                os.remove(path)
                
                with st.spinner('Sincronizando Banco de Dados...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    
                    # PROJETOS
                    try: ws_p = sh.worksheet("Projetos")
                    except: ws_p = sh.add_worksheet("Projetos", 100, 5)
                    recs = ws_p.get_all_records()
                    df = pd.DataFrame(recs)
                    if not df.empty and 'ID_Projeto' in df.columns: df = df[df['ID_Projeto'] != id_proj]
                    
                    # Calcula volume total da obra para salvar no projeto
                    vol_total = sum(d['Volume_Concreto_m3'] for d in dados)
                    
                    new = {
                        'ID_Projeto': id_proj, 'Nome_Obra': nome, 
                        'Data_Upload': datetime.datetime.now().strftime("%d/%m/%Y"), 
                        'Total_Pilares': len(dados),
                        'Volume_Total_Concreto': round(vol_total, 2)
                    }
                    df_final = pd.concat([df, pd.DataFrame([new])], ignore_index=True).fillna("").astype(str)
                    ws_p.clear()
                    ws_p.update([df_final.columns.values.tolist()] + df_final.values.tolist())

                    # PILARES
                    try: ws_pil = sh.worksheet("Pilares")
                    except: ws_pil = sh.add_worksheet("Pilares", 1000, 10)
                    recs = ws_pil.get_all_records()
                    df = pd.DataFrame(recs)
                    if not df.empty and 'Projeto_Ref' in df.columns: df = df[df['Projeto_Ref'] != id_proj]
                    
                    df_pil_final = pd.concat([df, pd.DataFrame(dados)], ignore_index=True)
                    df_pil_final = df_pil_final.fillna("").astype(str)
                    
                    ws_pil.clear()
                    ws_pil.update([df_pil_final.columns.values.tolist()] + df_pil_final.values.tolist())
                
                with st.spinner('Gerando PDF...'):
                    pdf = gerar_pdf_memoria(dados, nome)
                
                st.success(f"✅ Sucesso! {len(dados)} pilares processados.")
                st.info(f"📊 Volume Total de Concreto Estimado: {round(vol_total, 2)} m³")
                st.download_button("📥 BAIXAR PDF", pdf, f"Etiquetas_{id_proj}.pdf", "application/pdf")
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
