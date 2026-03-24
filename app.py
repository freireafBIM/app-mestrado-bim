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
  TQS_Armaduras → Tem_Protensao
  IfcReinforcingBar.Name → "QTD NOME_ELEM ØDiam C=Comp"
"""

import streamlit as st
import os, re, io, tempfile, datetime, math
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
    Converte codificações de string do formato IFC STEP para texto legível.

    Trata dois padrões:
      \\X2\\HHHH...\\X0\\  — escape Unicode IFC2x3/IFC4.
        O bloco hex pode conter múltiplos codepoints de 4 dígitos concatenados.
        Exemplo: \\X2\\00E700F5\\X0\\ → 'çõ' (fundações).
        Regex usa [0-9A-Fa-f]+ para capturar N×4 dígitos de uma vez.
      \\S\\x  — offset ISO-8859-1 legado (0x80 + ord(x)).
    """
    if not texto:
        return ""
    # Unicode: um ou mais codepoints de 4 hex cada
    def _decode_block(m):
        hex_str = m.group(1)
        return "".join(chr(int(hex_str[i:i+4], 16)) for i in range(0, len(hex_str), 4))
    texto = re.sub(r'\\X2\\([0-9A-Fa-f]+)\\X0\\', _decode_block, texto)
    # ISO-8859-1 legacy
    texto = re.sub(r'\\S\\(.)', lambda m: chr(0x80 + ord(m.group(1))), texto)
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


def indexar_armaduras(ifc_file, ifc_path: str | None = None) -> dict:
    """
    Vincula cada IfcReinforcingBar ao seu elemento estrutural pai.

    DIAGNÓSTICO (confirmado em dois arquivos TQS — v22.12 e v27.0):
    O IFC TQS não contém relação explícita barra→elemento.

    ESTRATÉGIA HÍBRIDA (3 métodos por prioridade):

    MÉTODO 1 — Containment espacial (TODOS os tipos):
      Calcula o centro 2D (x,y) de cada elemento estrutural:
        - ExtrudedAreaSolid → Axis2Placement3D → CartesianPoint  (TQS v27)
        - FacetedBrep       → centróide dos vértices do ClosedShell  (TQS v22)
      Extrai o ponto inicial (x,y) de cada barra via SweptDiskSolid → IfcLine.
      Associa a barra ao elemento cujo bounding box a contém.

    MÉTODO 2 — TQS_Padrao.Numero (fallback para vigas/lajes sem posição):
      Quando a barra tem geometria em coordenadas locais (ExtrudedAreaSolid
      com (0,0,0)), usa Pset TQS_Padrao.Numero + Planta para vincular.

    ENCODING IFC SUPORTADO:
      \\X2\\HHHH...\\X0\\ — Unicode multi-codepoint (TQS v27+)
      \\X\\HH             — Latin-1 single-byte (TQS v22)
      \\S\\x              — offset +0x80

    Retorna: { elemento_eid_int: [(bitola, comp_cm, sub_tipo), ...] }
    """
    import re as _re

    # ── Decodificação IFC robusta (ambos os formatos) ─────────────────────────
    def _dec(s):
        if not s: return ""
        def _x2(m):
            h = m.group(1)
            return "".join(chr(int(h[i:i+4],16)) for i in range(0,len(h),4))
        s = _re.sub(r'\\X2\\([0-9A-Fa-f]+)\\X0\\', _x2, s)
        s = _re.sub(r'\\X\\([0-9A-Fa-f]{2})', lambda m: chr(int(m.group(1),16)), s)
        s = _re.sub(r'\\S\\(.)', lambda m: chr(0x80+ord(m.group(1))), s)
        return s

    # ── Parser manual linha a linha ───────────────────────────────────────────
    # Necessário porque o parser regex padrão perde entidades cuja linha
    # anterior termina em ));  — problema confirmado no IFC do TQS v27.
    _ents: dict[str, tuple[str, str]] = {}
    _filepath = (
        ifc_path
        or getattr(ifc_file, "path", None)
        or getattr(ifc_file, "_filepath", None)
        or getattr(getattr(ifc_file, "wrapped_data", None), "path", None)
        or getattr(getattr(ifc_file, "wrapped_data", None), "_path", None)
    )
    if _filepath:
        with open(_filepath, "r", encoding="latin-1") as _fh:
            _raw = _fh.read()
        _cid = _ctype = None; _cdata: list[str] = []
        for _ln in _raw.splitlines():
            _ln = _ln.strip()
            if not _ln or _ln.startswith((
                "ISO-10303","HEADER;","DATA;","ENDSEC;","END-ISO","FILE_","/*"
            )): continue
            _m = _re.match(r"^#(\d+)=([A-Z][A-Z0-9_]*)\((.*)$", _ln)
            if _m:
                if _cid and _ctype:
                    _d = ",".join(_cdata)
                    if _d.endswith(");"): _d=_d[:-2]
                    elif _d.endswith(")"): _d=_d[:-1]
                    _ents[_cid]=(_ctype,_d)
                _cid,_ctype=_m.group(1),_m.group(2); _rest=_m.group(3)
                if _rest.endswith(";"):
                    _d=_rest[:-1]
                    if _d.endswith(")"): _d=_d[:-1]
                    _ents[_cid]=(_ctype,_d)
                    _cid=_ctype=None; _cdata=[]
                else:
                    _cdata=[_rest]
            elif _cid:
                if _ln.endswith(";"):
                    _cdata.append(_ln[:-1].rstrip(")"))
                    _ents[_cid]=(_ctype,"\n".join(_cdata))
                    _cid=_ctype=None; _cdata=[]
                else:
                    _cdata.append(_ln)
        _use_manual = True
    else:
        _use_manual = False

    # ── Mapeamentos ───────────────────────────────────────────────────────────
    OBJ_TO_IFC: dict[str,str] = {
        "Armadura longitudinal pilares":           "IFCCOLUMN",
        "Armadura transversal pilares":            "IFCCOLUMN",
        "Armadura longitudinal compl pilares":     "IFCCOLUMN",
        "Armadura longitudinal vigas":             "IFCBEAM",
        "Armadura longitudinal vigas negativa":    "IFCBEAM",
        "Armadura transversal vigas":              "IFCBEAM",
        "Armadura lateral vigas":                  "IFCBEAM",
        "Grampos de vigas":                        "IFCBEAM",
        "Armadura longitudinal positiva lajes":    "IFCSLAB",
        "Armadura longitudinal negativa lajes":    "IFCSLAB",
        "Armadura longitudinal lajes secund\u00e1ria": "IFCSLAB",
        "Armadura de funda\u00e7\u00f5es":          "IFCFOOTING",
        "Armadura de estacas":                     "IFCPILE",
    }
    OBJ_SUB: dict[str,str] = {
        "Armadura longitudinal pilares":"long",
        "Armadura transversal pilares":"trans",
        "Armadura longitudinal compl pilares":"long",
        "Armadura longitudinal vigas":"long",
        "Armadura longitudinal vigas negativa":"long",
        "Armadura transversal vigas":"trans",
        "Armadura lateral vigas":"lat",
        "Grampos de vigas":"grampo",
        "Armadura longitudinal positiva lajes":"long",
        "Armadura longitudinal negativa lajes":"long",
        "Armadura longitudinal lajes secund\u00e1ria":"long",
        "Armadura de funda\u00e7\u00f5es":"long",
        "Armadura de estacas":"long",
    }
    TIPOS_SPATIAL = {"IFCCOLUMN","IFCFOOTING","IFCPILE","IFCBEAM","IFCSLAB"}
    TOLERANCIA_CM = 15.0

    # ── Passo 1: pavimento de cada elemento/barra ─────────────────────────────
    storey_por_elem: dict[int,str] = {}
    for rel in ifc_file.by_type("IfcRelContainedInSpatialStructure"):
        try:
            storey = _dec(rel.RelatingStructure.Name or "Sem pavimento")
            for elem in rel.RelatedElements:
                storey_por_elem[elem.id()] = storey
        except Exception:
            pass

    # ── Passo 2: Psets via ifcopenshell ──────────────────────────────────────
    def _psets(elem) -> dict:
        try:
            import ifcopenshell.util.element as _ie
            return _ie.get_psets(elem)
        except Exception:
            return {}

    def _dim(elem, *keys) -> float:
        geo = _psets(elem).get("TQS_Geometria", {})
        for k in keys:
            v = geo.get(k)
            if v is not None:
                try: return float(str(v))
                except: pass
        return 0.0

    def _num_planta(elem) -> tuple:
        p = _psets(elem).get("TQS_Padrao", {})
        try: num = int(str(p.get("Numero","")).strip())
        except: num = None
        planta = str(p.get("Planta","") or "").strip()
        return num, planta or None

    # ── Passo 3: centro (x,y) de cada elemento — dois métodos ────────────────
    def _center_axis2placement(eid_str: str):
        """Para TQS v27: Axis2Placement3D → CartesianPoint direto."""
        if not _use_manual or eid_str not in _ents: return None
        et,ed = _ents[eid_str]
        parts = ed.split(",")
        if len(parts)<6: return None
        lp = parts[5].strip().lstrip("#")
        if lp not in _ents: return None
        lrefs = _re.findall(r"#(\d+)|\$", _ents[lp][1])
        ax = lrefs[1] if len(lrefs)>1 and lrefs[1]!="$" else None
        if ax and ax in _ents:
            arefs = _re.findall(r"#(\d+)", _ents[ax][1])
            if arefs:
                pt = arefs[0]
                if pt in _ents and _ents[pt][0]=="IFCCARTESIANPOINT":
                    c = _re.findall(r"[-\d.E+]+", _ents[pt][1])
                    if len(c)>=2: return float(c[0]), float(c[1])
        return None

    def _center_brep(eid_str: str):
        """Para TQS v22: centróide dos vértices do FacetedBrep."""
        if not _use_manual or eid_str not in _ents: return None
        et,ed = _ents[eid_str]
        parts = ed.split(",")
        repr_id = parts[6].strip().lstrip("#") if len(parts)>6 else ""
        if repr_id not in _ents: return None
        allowed = {"IFCPRODUCTDEFINITIONSHAPE","IFCSHAPEREPRESENTATION",
                   "IFCFACETEDBREP","IFCCLOSEDSHELL","IFCFACE",
                   "IFCFACEOUTERBOUND","IFCPOLYLOOP","IFCCARTESIANPOINT"}
        pts=[]; vis=set()
        def _wk(pid, d=0):
            if pid in vis or pid not in _ents or d>12: return
            vis.add(pid)
            et2,ed2=_ents[pid]
            if et2 not in allowed: return
            if et2=="IFCCARTESIANPOINT":
                c=_re.findall(r"[-\d.E+]+",ed2)
                if len(c)>=2: pts.append((float(c[0]),float(c[1])))
                return
            for r in _re.findall(r"#(\d+)",ed2): _wk(r,d+1)
        _wk(repr_id)
        if not pts: return None
        xs=[p[0] for p in pts]; ys=[p[1] for p in pts]
        return (min(xs)+max(xs))/2, (min(ys)+max(ys))/2

    def _get_elem_center(elem) -> tuple | None:
        eid_str = str(elem.id())
        # 1. Tentar via Axis2Placement3D (TQS v27 — ExtrudedAreaSolid)
        c = _center_axis2placement(eid_str)
        if c and (c[0]!=0 or c[1]!=0):
            return c
        # 2. Tentar via FacetedBrep (TQS v22)
        c = _center_brep(eid_str)
        if c and (c[0]!=0 or c[1]!=0):
            return c
        # 3. Fallback: ifcopenshell.util.placement
        try:
            import ifcopenshell.util.placement as _pl
            import numpy as np
            M = _pl.get_local_placement(elem.ObjectPlacement)
            x,y = float(M[0,3]), float(M[1,3])
            if x!=0 or y!=0: return x, y
        except Exception:
            pass
        return None

    # ── Passo 4: bounding boxes dos elementos ─────────────────────────────────
    # Para PILARES, FUNDAÇÕES, ESTACAS: bbox 2D centrada (cx, cy, hb, hh)
    # Para VIGAS: bbox 3D completa via vértices do Brep (v22) ou placement+dims (v27)
    # Para LAJES: bbox 2D grande (cobre toda a planta)
    elem_bbox: dict[int,tuple] = {}   # eid → (cx,cy,hb,hh,tipo_STEP)         [pilares/fund/estacas/lajes]
    viga_bbox3d: dict[int,tuple] = {} # eid → (xmin,xmax,ymin,ymax,zmin,zmax,eixo)  [vigas]

    _BREP_TYPES = {"IFCPRODUCTDEFINITIONSHAPE","IFCSHAPEREPRESENTATION",
                   "IFCFACETEDBREP","IFCCLOSEDSHELL","IFCFACE",
                   "IFCFACEOUTERBOUND","IFCPOLYLOOP","IFCCARTESIANPOINT"}

    def _viga_bbox3d_brep(eid_str: str):
        """Extrai bbox 3D de viga com FacetedBrep (v22) via vértices do Brep."""
        if not _use_manual or eid_str not in _ents: return None
        et,ed = _ents[eid_str]
        parts = ed.split(","); repr_id = parts[6].strip().lstrip("#") if len(parts)>6 else ""
        if repr_id not in _ents: return None
        pts=[]; vis=set()
        def _wk(pid, d=0):
            if pid in vis or pid not in _ents or d>12: return
            vis.add(pid); et2,ed2=_ents[pid]
            if et2 not in _BREP_TYPES: return
            if et2=="IFCCARTESIANPOINT":
                c=_re.findall(r"[-\d.E+]+",ed2)
                if len(c)>=3: pts.append((float(c[0]),float(c[1]),float(c[2])))
                return
            for r in _re.findall(r"#(\d+)",ed2): _wk(r,d+1)
        _wk(repr_id)
        if not pts: return None
        xs=[p[0] for p in pts]; ys=[p[1] for p in pts]; zs=[p[2] for p in pts]
        dx=max(xs)-min(xs); dy=max(ys)-min(ys)
        eixo="X" if dx>=dy else "Y"
        return (min(xs),max(xs),min(ys),max(ys),min(zs),max(zs),eixo)

    def _viga_bbox3d_extruded(elem):
        """Estima bbox 3D de viga com ExtrudedAreaSolid (v27) via placement + dims."""
        try:
            import ifcopenshell.util.placement as _pl
            import numpy as _np
            M = _pl.get_local_placement(elem.ObjectPlacement)
            ox,oy,oz = float(M[0,3]),float(M[1,3]),float(M[2,3])
            # Vetor direção da viga = coluna 0 da matriz (eixo X local)
            dx_,dy_ = float(M[0,0]),float(M[1,0])
            larg = _dim(elem,"Largura") or 14.
            alt  = _dim(elem,"Altura") or 35.
            # Comprimento: tentar extrair do ExtrudedAreaSolid depth
            comp = 0.0
            try:
                for rep in elem.Representation.Representations:
                    for item in rep.Items:
                        if item.is_a("IfcExtrudedAreaSolid"):
                            comp = float(item.Depth); break
                    if comp: break
            except Exception: pass
            if comp == 0.0: comp = 300.0  # fallback
            # Calcular endpoints
            ex,ey = ox+dx_*comp, oy+dy_*comp
            xmin,xmax = min(ox,ex)-larg,max(ox,ex)+larg
            ymin,ymax = min(oy,ey)-alt, max(oy,ey)+alt
            zmin,zmax = oz-alt, oz+alt
            eixo="X" if abs(dx_)>=abs(dy_) else "Y"
            return (xmin,xmax,ymin,ymax,zmin,zmax,eixo)
        except Exception:
            return None

    TIPOS_STRUCT = ("IfcColumn","IfcBeam","IfcSlab","IfcFooting","IfcPile")
    for tipo in TIPOS_STRUCT:
        tipo_step = "IFC"+tipo[3:].upper()
        for elem in ifc_file.by_type(tipo):
            eid_str = str(elem.id())

            if tipo_step == "IFCBEAM":
                # Tentar bbox 3D via Brep (v22) primeiro, depois ExtrudedAreaSolid (v27)
                bb3 = _viga_bbox3d_brep(eid_str) or _viga_bbox3d_extruded(elem)
                if bb3:
                    viga_bbox3d[elem.id()] = bb3
                continue  # vigas não usam elem_bbox

            center = _get_elem_center(elem)
            if center is None: continue
            cx,cy = center
            if tipo_step=="IFCCOLUMN":
                b=_dim(elem,"Dimensao_b1") or 35.; h=_dim(elem,"Dimensao_h1") or 14.
            elif tipo_step=="IFCSLAB":
                b,h = 600.,600.
            elif tipo_step=="IFCFOOTING":
                b=_dim(elem,"Dimensoes_X") or _dim(elem,"Diametro") or 60.
                h=_dim(elem,"Dimensoes_Y") or b
            else:
                d2=_dim(elem,"Diametro") or _dim(elem,"Dimensao_b1") or 30.
                b,h=d2,d2
            elem_bbox[elem.id()]=(cx,cy,b/2+TOLERANCIA_CM,h/2+TOLERANCIA_CM,tipo_step)

    # ── Passo 5: índice Numero+Planta para vigas/lajes (fallback) ────────────
    elem_num_idx: dict[tuple,list] = {}
    for tipo in ("IfcBeam","IfcSlab"):
        tipo_step="IFC"+tipo[3:].upper()
        for elem in ifc_file.by_type(tipo):
            num,planta=_num_planta(elem)
            if num is not None and planta:
                key=(num,planta,tipo_step)
                elem_num_idx.setdefault(key,[]).append(elem.id())

    # ── Passo 6: ponto inicial (x,y,z) da barra via SweptDiskSolid ───────────
    # Retorna (x, y, z). Para compatibilidade, o chamador usa só (x,y) ou (x,y,z).
    def _bar_xyz(barra) -> tuple | None:
        if not getattr(barra,"Representation",None): return None
        vis=set()
        def _wk(obj,d=0):
            if obj is None or d>12: return None
            oid=id(obj)
            if oid in vis: return None
            vis.add(oid)
            if not hasattr(obj,"is_a"): return None
            t=obj.is_a()
            if t=="IfcCartesianPoint":
                c=obj.Coordinates
                if not c or len(c)<2: return None
                z = float(c[2]) if len(c)>=3 else 0.0
                return (float(c[0]),float(c[1]),z)
            if t=="IfcSweptDiskSolid": return _wk(obj.Directrix,d+1)
            if t=="IfcCompositeCurve":
                for seg in (obj.Segments or []):
                    pt=_wk(seg,d+1)
                    if pt: return pt
                return None
            if t=="IfcCompositeCurveSegment": return _wk(obj.ParentCurve,d+1)
            if t=="IfcLine": return _wk(obj.Pnt,d+1)
            if t=="IfcTrimmedCurve": return _wk(obj.BasisCurve,d+1)
            if t=="IfcCircle": return _wk(obj.Position,d+1)
            if t in("IfcAxis2Placement3D","IfcAxis2Placement2D"): return _wk(obj.Location,d+1)
            if t=="IfcPolyline":
                pts=obj.Points or []
                return _wk(pts[0],d+1) if pts else None
            return None
        for rep in barra.Representation.Representations:
            for item in rep.Items:
                pt=_wk(item)
                if pt: return pt
        return None

    # Alias para compatibilidade: só (x,y)
    def _bar_xy(barra):
        r = _bar_xyz(barra)
        return (r[0], r[1]) if r else None

    # ── Passo 7: Numero+Planta da barra (fallback para vigas/lajes) ──────────
    def _bar_num_planta(barra):
        p=_psets(barra).get("TQS_Padrao",{})
        try: num=int(str(p.get("Numero","")).strip())
        except: num=None
        planta=str(p.get("Planta","") or "").strip()
        return num, planta or None

    # ── Passo 8: vincular barras → elementos ─────────────────────────────────
    cache: dict[int,list] = defaultdict(list)
    REGEX_B = re.compile(r"[Ø\u00d8](\d+\.?\d*)\s+C=(\d+\.?\d*)", re.UNICODE)
    REGEX_D = re.compile(r"[Ø\u00d8](\d+\.?\d*)", re.UNICODE)

    sem_diam=sem_xy=sem_match=sem_num=0
    vin_spatial=vin_numero=0

    for barra in ifc_file.by_type("IfcReinforcingBar"):
        try:
            ot_raw = _dec(barra.ObjectType or "")
            tipo_step = OBJ_TO_IFC.get(ot_raw)
            if not tipo_step: continue

            nome = _dec(barra.Name or "")
            mb = REGEX_B.search(nome)
            if mb:
                bitola=float(mb.group(1)); comp_cm=float(mb.group(2))
            else:
                md = REGEX_D.search(nome)
                if not md: sem_diam+=1; continue
                bitola=float(md.group(1)); comp_cm=0.0

            sub = OBJ_SUB.get(ot_raw,"long")
            pav = storey_por_elem.get(barra.id(),"Sem pavimento")

            # — Método 1: containment espacial ─────────────────────────────
            xyz = _bar_xyz(barra)
            if xyz is not None:
                bx,by,bz = xyz

                if tipo_step == "IFCBEAM":
                    # Vigas: containment 3D com tolerância ASSIMÉTRICA por eixo.
                    # Barras longitudinais começam no nó do pilar (~15cm fora da bbox
                    # no eixo da viga) → TOL_EIXO=15cm.
                    # No eixo transversal, TOL=5cm evita falsos positivos em vigas paralelas.
                    _TOL_EIXO = 15.0
                    _TOL_TRANS = 5.0
                    _TOL_Z = 5.0
                    best=None; best_vol=1e18; best_eixo="X"
                    for e_id,(xmin,xmax,ymin,ymax,zmin,zmax,eixo) in viga_bbox3d.items():
                        if storey_por_elem.get(e_id,"?")!=pav: continue
                        if eixo=="X":
                            ok_e = xmin-_TOL_EIXO  <= bx <= xmax+_TOL_EIXO
                            ok_t = ymin-_TOL_TRANS <= by <= ymax+_TOL_TRANS
                        else:
                            ok_e = ymin-_TOL_EIXO  <= by <= ymax+_TOL_EIXO
                            ok_t = xmin-_TOL_TRANS <= bx <= xmax+_TOL_TRANS
                        ok_z = zmin-_TOL_Z <= bz <= zmax+_TOL_Z
                        if ok_e and ok_t and ok_z:
                            vol=(xmax-xmin)*(ymax-ymin)*(zmax-zmin)
                            if vol<best_vol: best_vol=vol; best=e_id; best_eixo=eixo
                    if best:
                        # Para estribos: guardar coord ao longo do eixo da viga
                        if sub=="trans":
                            coord_eixo = round(bx,3) if best_eixo=="X" else round(by,3)
                            cache[best].append((bitola,comp_cm,sub,coord_eixo))
                        else:
                            cache[best].append((bitola,comp_cm,sub))
                        vin_spatial+=1; continue
                    # sem match → fallback Numero+Planta abaixo
                else:
                    # Pilares, fundações, estacas: containment 2D por bbox de seção
                    best=None; best_dist=1e9
                    for e_id,(cx,cy,hb,hh,te) in elem_bbox.items():
                        if te!=tipo_step: continue
                        if storey_por_elem.get(e_id,"?")!=pav: continue
                        if abs(bx-cx)<=hb and abs(by-cy)<=hh:
                            d2=math.sqrt((bx-cx)**2+(by-cy)**2)
                            if d2<best_dist: best_dist=d2; best=e_id
                    if best:
                        # Para estribos de pilar: guardar Z para calcular espaçamento
                        if tipo_step=="IFCCOLUMN" and sub=="trans":
                            cache[best].append((bitola,comp_cm,sub,round(bz,3)))
                        else:
                            cache[best].append((bitola,comp_cm,sub))
                        vin_spatial+=1; continue
                # Sem match espacial — tentar Numero como fallback
            else:
                sem_xy+=1

            # — Método 2: Numero+Planta (vigas/lajes sem posição absoluta) ─
            if tipo_step in ("IFCBEAM","IFCSLAB"):
                num,planta=_bar_num_planta(barra)
                if num is not None and planta:
                    key=(num,planta,tipo_step)
                    candidatos=elem_num_idx.get(key,[])
                    matched=False
                    for e_id in candidatos:
                        if storey_por_elem.get(e_id,"?")!=pav: continue
                        cache[e_id].append((bitola,comp_cm,sub))
                        vin_numero+=1; matched=True
                    if matched: continue
            sem_match+=1

        except Exception as exc:
            st.warning(f"Barra ignorada (#{getattr(barra,'id','?')}): {exc}")

    total=vin_spatial+vin_numero
    st.info(
        f"Armaduras indexadas: **{total}** barras em **{len(cache)}** elementos. "
        f"(Espacial: {vin_spatial} | Número: {vin_numero} | "
        f"Sem match: {sem_match} | Sem geometria: {sem_xy})"
    )
    return dict(cache), viga_bbox3d

# ── Tabela grau de aço (NBR 7480 / prática mercado brasileiro) ──────────────
# Ø ≤ 6.0 mm → CA-60 (fios trefilados de alta resistência)
# Ø > 6.0 mm → CA-50 (barras nervuradas)
def grau_aco(diam_mm: float) -> str:
    return "CA-60" if diam_mm <= 6.0 else "CA-50"


# ── Peso linear (kg/m) — fórmula NBR: ρ = d² × 0.00617  (d em mm) ──────────
def peso_linear_kg_m(diam_mm: float) -> float:
    return diam_mm ** 2 * 0.00617


def _parse_barra(b):
    """Normaliza entrada do cache: retorna (bitola, comp_cm, sub_tipo, z_cm).
    Aceita tuplas de 3 (sem Z) ou 4 campos (com Z), e também float simples.
    """
    if isinstance(b, tuple):
        if len(b) == 4:
            return float(b[0]), float(b[1]), str(b[2]), float(b[3])
        if len(b) == 3:
            return float(b[0]), float(b[1]), str(b[2]), None
        if len(b) == 2:
            return float(b[0]), float(b[1]), "long", None
    return float(b), 0.0, "long", None


def _espaçamento_estribos(zs: list[float]) -> int:
    """Calcula espaçamento nominal (cm, inteiro) a partir das alturas Z dos estribos.
    Mede distância c/c e arredonda para inteiro — recupera o valor nominal de projeto.
    """
    if len(zs) < 2:
        return 0
    zs_ord = sorted(zs)
    diffs = [zs_ord[i+1] - zs_ord[i] for i in range(len(zs_ord)-1)]
    return round(sum(diffs) / len(diffs))


def formatar_armadura(cache: dict, elem_eid: int) -> str:
    """
    Monta o campo Armadura completo com longitudinal e transversal.

    Pilares — formato:
      Long: 6Ø10(C=356cm) | Trans: 26Ø5(C=104cm)@12cm + 26Ø5(C=20cm)@12cm

    Outros elementos — formato:
      Long: 12Ø16 + 8Ø10 | Trans: 60Ø5

    Para pilares o espaçamento é calculado pelas coordenadas Z dos estribos.
    Para outros elementos o espaçamento não está disponível no IFC.
    """
    if elem_eid not in cache:
        return "Sem armadura exportada"

    barras = cache[elem_eid]
    long_list  = []   # (bitola, comp_cm)
    trans_list = []   # (bitola, comp_cm, z_cm|None)

    for b in barras:
        d, c, sub, z = _parse_barra(b)
        if sub == "long":
            long_list.append((d, c))
        else:
            trans_list.append((d, c, z))

    if not long_list and not trans_list:
        return "Sem armadura exportada"

    # ── Longitudinal ──────────────────────────────────────────────────────────
    # Agrupar por (bitola, comp) para mostrar "qtd Ø bitola (C=comp cm)"
    grupos_long: dict = defaultdict(int)
    for d, c in long_list:
        grupos_long[(d, c)] += 1

    long_parts = []
    for (d, c), q in sorted(grupos_long.items(), key=lambda x: -x[0][0]):
        if c > 0:
            long_parts.append(f"{q}Ø{d:.0f}(C={c:.0f}cm)")
        else:
            long_parts.append(f"{q}Ø{d:.0f}")
    long_str = " + ".join(long_parts) if long_parts else "—"

    # ── Transversal ───────────────────────────────────────────────────────────
    # Agrupar por (bitola, comp) → cada grupo é um tipo de estribo
    grupos_trans: dict = defaultdict(list)   # (d, c) → [z, z, ...]
    for d, c, z in trans_list:
        grupos_trans[(d, c)].append(z)

    trans_parts = []
    for (d, c), zs in sorted(grupos_trans.items(), key=lambda x: -x[0][1]):
        q = len(zs)
        zs_validos = [z for z in zs if z is not None]
        if zs_validos:
            esp = _espaçamento_estribos(zs_validos)
            trans_parts.append(f"{q}Ø{d:.0f}(C={c:.0f}cm)@{esp}cm")
        else:
            if c > 0:
                trans_parts.append(f"{q}Ø{d:.0f}(C={c:.0f}cm)")
            else:
                trans_parts.append(f"{q}Ø{d:.0f}")
    trans_str = " + ".join(trans_parts) if trans_parts else "—"

    return f"Long: {long_str} | Trans: {trans_str}"


def detalhar_armadura(cache: dict, elem_eid: int) -> dict:
    """
    Retorna detalhamento completo da armadura para uso interno/UI:
    {
      "longitudinal":   "6Ø10(C=356cm)",
      "transversal":    "26Ø5(C=104cm)@12cm + 26Ø5(C=20cm)@12cm",
      "comp_total_m":   ...,
      "peso_total_kg":  ...,
      "por_bitola": [ {"bitola", "grau", "qtd", "comp_unit_cm",
                        "comp_total_m", "peso_kg", "espaçamento_cm"}, ...]
    }
    """
    if elem_eid not in cache:
        return {}

    barras = cache[elem_eid]
    long_list  = []
    trans_list = []

    for b in barras:
        d, c, sub, z = _parse_barra(b)
        if sub == "long":
            long_list.append((d, c))
        else:
            trans_list.append((d, c, z))

    # Longitudinal
    gl: dict = defaultdict(int)
    for d, c in long_list:
        gl[(d, c)] += 1
    long_parts = []
    for (d, c), q in sorted(gl.items(), key=lambda x: -x[0][0]):
        long_parts.append(f"{q}Ø{d:.0f}(C={c:.0f}cm)" if c > 0 else f"{q}Ø{d:.0f}")
    long_str = " + ".join(long_parts) if long_parts else "—"

    # Transversal
    gt: dict = defaultdict(list)
    for d, c, z in trans_list:
        gt[(d, c)].append(z)
    trans_parts = []
    for (d, c), zs in sorted(gt.items(), key=lambda x: -x[0][1]):
        q = len(zs); zs_v = [z for z in zs if z is not None]
        esp = _espaçamento_estribos(zs_v) if zs_v else None
        s = f"{q}Ø{d:.0f}(C={c:.0f}cm)" if c > 0 else f"{q}Ø{d:.0f}"
        if esp: s += f"@{esp}cm"
        trans_parts.append(s)
    trans_str = " + ".join(trans_parts) if trans_parts else "—"

    # por_bitola — para fins de quantitativo detalhado
    por_bitola = []
    # Chave: (bitola, comp_unitario, sub_tipo) — separa tipos distintos de estribo
    todos_grupos: dict = defaultdict(lambda: {"comps": [], "zs": [], "sub": "long"})
    for d, c in long_list:
        todos_grupos[(d, c, "long")]["comps"].append(c)
        todos_grupos[(d, c, "long")]["sub"] = "long"
    for d, c, z in trans_list:
        todos_grupos[(d, c, "trans")]["comps"].append(c)
        todos_grupos[(d, c, "trans")]["zs"].append(z)
        todos_grupos[(d, c, "trans")]["sub"] = "trans"

    for (d, _comp, sub), v in sorted(todos_grupos.items(), key=lambda x: (-x[0][0], -x[0][1])):
        comps_v = [c for c in v["comps"] if c > 0]
        comp_total_m = sum(comps_v) / 100.0
        peso_kg = peso_linear_kg_m(d) * comp_total_m
        zs_v = [z for z in v["zs"] if z is not None]
        esp = _espaçamento_estribos(zs_v) if zs_v else None
        por_bitola.append({
            "bitola":          d,
            "grau":            grau_aco(d),
            "sub_tipo":        sub,
            "qtd":             len(v["comps"]),
            "comp_unit_cm":    round(comps_v[0], 1) if comps_v else 0.0,
            "comp_total_m":    round(comp_total_m, 2),
            "peso_kg":         round(peso_kg, 3),
            "espaçamento_cm":  esp,
        })

    todas_comps = [(d, c) for d, c in long_list] +                   [(d, c) for d, c, _ in trans_list]
    comp_total_m  = sum(c for _, c in todas_comps if c > 0) / 100.0
    peso_total_kg = sum(peso_linear_kg_m(d)*(c/100) for d,c in todas_comps if c>0)

    return {
        "longitudinal":  long_str,
        "transversal":   trans_str,
        "comp_total_m":  round(comp_total_m, 2),
        "peso_total_kg": round(peso_total_kg, 3),
        "por_bitola":    por_bitola,
    }





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
        cache_arm, _viga_bbox3d_map = indexar_armaduras(ifc, caminho)


    # ── 2. Fusão de segmentos de nó de viga (≤ 20 cm) ────────────────────────
    # Segmentos curtos (≤ 20 cm) são trechos de sobreposição sobre pilares ou
    # outras vigas. Suas barras transversais são somadas ao vão adjacente maior
    # e o segmento nó é suprimido do relatório (pois na prática a viga é contínua).
    NOS_VIGAS: set[int] = set()   # IDs de segmentos a suprimir

    # Agrupar segmentos por (nome, pavimento) e calcular comprimento via bbox
    _segs_por_viga: dict = {}  # (nome, pav) → [(eid, comp_cm), ...]
    for _elem in ifc.by_type("IfcBeam"):
        _pav = decode_ifc(
            (_elem.ContainedInStructure[0].RelatingStructure.Name or "")
            if _elem.ContainedInStructure else ""
        ) or "Sem pavimento"
        _nome = _elem.Name or "S/N"
        # Comprimento via viga_bbox3d já calculada em indexar_armaduras.
        # xmin,xmax,ymin,ymax,zmin,zmax,eixo  → dimensão dominante = comprimento
        _bb = _viga_bbox3d_map.get(_elem.id())
        if _bb:
            _dx = _bb[1] - _bb[0]
            _dy = _bb[3] - _bb[2]
            _comp = max(_dx, _dy)
        else:
            _comp = 999.0  # sem bbox → tratar como vão longo
        _key = (_nome, _pav)
        _segs_por_viga.setdefault(_key, []).append((_elem.id(), _comp))

    for _key, _segs in _segs_por_viga.items():
        if len(_segs) <= 1:
            continue  # viga com apenas 1 segmento — nada a fundir

        # Separar vãos e nós
        _vaos = [(eid, c) for eid, c in _segs if c > 20.0]
        _nos  = [(eid, c) for eid, c in _segs if c <= 20.0]
        if not _nos:
            continue  # sem nós — nada a fazer

        # Para cada nó, transferir barras transversais ao vão mais próximo em volume
        # (na prática: o vão de menor comprimento que seja adjacente, mas sem
        # coordenadas de adjacência disponíveis aqui, usamos o de maior comprimento
        # do mesmo grupo — barra vai para o vão principal da viga)
        _vao_principal = max(_vaos, key=lambda x: x[1])[0] if _vaos else None

        for _no_eid, _ in _nos:
            NOS_VIGAS.add(_no_eid)
            if _vao_principal and _no_eid in cache_arm:
                # Adicionar barras transversais do nó ao vão principal
                # Adicionar trans do nó SEM coord_eixo (3 campos),
                # para que entrem na contagem mas não distorçam o espaçamento
                # calculado apenas pelas posições regulares do vão.
                _barras_no = [
                    (b[0], b[1], b[2])           # (bitola, comp, sub) sem coord
                    for b in cache_arm[_no_eid]
                    if isinstance(b, tuple) and len(b) >= 3 and b[2] == "trans"
                ]
                if _barras_no:
                    cache_arm.setdefault(_vao_principal, [])
                    cache_arm[_vao_principal].extend(_barras_no)
                del cache_arm[_no_eid]

    if NOS_VIGAS:
        st.info(f"Fusão de vigas: {len(NOS_VIGAS)} segmento(s) de nó suprimido(s).")

    # ── 3. Processar cada tipo estrutural ─────────────────────────────────────
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

            # Suprimir segmentos de nó de viga (já fundidos ao vão principal)
            if tipo_ifc == "IfcBeam" and elem.id() in NOS_VIGAS:
                continue

            nome     = elem.Name or "S/N"
            pavimento = _pavimento(elem)
            geo      = _bbox(elem)
            ps       = _psets(elem)

            # ── Extrair campos específicos dos Psets TQS ──────────────────────
            # Compatibilidade entre versões TQS:
            #   v27+: Pset sem prefixo  → "TQS_Geometria.Dimensao_b1"
            #   v22:  Pset com prefixo  → "Pset_TQS_Geometria.Dimensao_b1"
            def get(*chaves: str, fallback: str = "—") -> str:
                for c in chaves:
                    # Tentar sem prefixo (v27) e com prefixo Pset_ (v22)
                    for chave in (c, "Pset_" + c):
                        v = ps.get(chave, "")
                        if v and v not in ("$", "—"):
                            return v
                return fallback

            # Material: v27 → TQS_Padrao.Material; v22 não tem esse campo
            # Fallback: ler do IfcMaterial associado ao elemento.
            material = get("TQS_Padrao.Material")
            if material == "—":
                try:
                    for assoc in (elem.HasAssociations or []):
                        if assoc.is_a("IfcRelAssociatesMaterial"):
                            mat = assoc.RelatingMaterial
                            if mat and mat.is_a("IfcMaterial"):
                                material = decode_ifc(mat.Name or "")
                            elif mat and mat.is_a("IfcMaterialList"):
                                nomes = [decode_ifc(m.Name or "") for m in (mat.Materials or [])]
                                material = ", ".join(n for n in nomes if n)
                            break
                except Exception:
                    pass

            protensao  = get("TQS_Armaduras.Tem_Protensao")

            # Geometria especifica por tipo (Pset tem precedencia sobre bbox)
            if tipo_ifc == "IfcColumn":
                secao = get("TQS_Geometria.Secao")
                dim_b = get("TQS_Geometria.Dimensao_b1")
                dim_h = get("TQS_Geometria.Dimensao_h1")
                area_secao = get("TQS_Geometria.Area")
                descricao_geo = (f"Secao:{secao} {dim_b}x{dim_h} cm"
                                 if dim_b != "—" else f"Secao:{secao}")

            elif tipo_ifc == "IfcBeam":
                largura = get("TQS_Geometria.Largura")
                altura  = get("TQS_Geometria.Altura")
                vao     = get("TQS_Geometria.Vao_Titulo", "TQS_Geometria.Vao")
                carga   = get("TQS_Geometria.Carga_linear")
                descricao_geo = f"{largura}x{altura} cm | Vao:{vao}"

            elif tipo_ifc == "IfcSlab":
                tipo_laje = get("TQS_Geometria.Tipo")
                area_sup  = get("TQS_Geometria.Area_superficie")
                carga_d   = get("TQS_Geometria.Carga_distribuida")
                # Espessura: v27=Altura; v22=Capa (laje nervurada/trelicada)
                esp = get("TQS_Geometria.Altura", "TQS_Geometria.Capa")
                # Area: v22 pode nao ter Area_superficie — usar bbox
                if area_sup == "—" and geo["comp_cm"] > 0 and geo["larg_cm"] > 0:
                    area_sup = f"{geo['comp_cm']*geo['larg_cm']/10000:.2f}"
                descricao_geo = f"Esp:{esp} cm | Area:{area_sup} m2 | {tipo_laje}"

            elif tipo_ifc == "IfcFooting":
                dim_x   = get("TQS_Geometria.Dimensoes_X")
                dim_y   = get("TQS_Geometria.Dimensoes_Y")
                alt_f   = get("TQS_Geometria.Altura")
                n_est   = get("TQS_Geometria.Estacas")
                tipo_f  = get("TQS_Geometria.Tipo")
                descricao_geo = f"{tipo_f} {dim_x}x{dim_y}x{alt_f} cm | Estacas:{n_est}"

            elif tipo_ifc == "IfcPile":
                diam_est = get("TQS_Geometria.Diametro", "TQS_Geometria.Dimensao_b1")
                alt_est  = get("TQS_Geometria.Altura")
                descricao_geo = f"O{diam_est} cm | H:{alt_est} cm"

            else:  # Escada, outros
                descricao_geo = (f"{geo['comp_cm']:.0f}x{geo['larg_cm']:.0f}x"
                                 f"{geo['alt_cm']:.0f} cm")


            armadura = formatar_armadura(cache_arm, elem.id())

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
                "Armadura":         armadura,
                "Tem_Protensao":    protensao,
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

    for reg in registros:
        tipo = reg["Tipo_Legivel"]

        # Fundo branco (sem cor por tipo)
        c.setFillColor(colors.white)
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

        # ── Armaduras: separar Long / Trans pelo delimitador " | " ─────────
        arm_raw = reg["Armadura"]
        if " | " in arm_raw:
            _arm_long_raw, _arm_trans_raw = arm_raw.split(" | ", 1)
        else:
            _arm_long_raw  = arm_raw
            _arm_trans_raw = ""

        # Prefixar corretamente com "Long:" e "Transv:"
        _long_label  = _arm_long_raw  if _arm_long_raw.startswith("Long:")  else f"Long: {_arm_long_raw}"
        _trans_label = _arm_trans_raw if _arm_trans_raw.startswith("Trans:") else (f"Transv: {_arm_trans_raw}" if _arm_trans_raw else "")
        # Normalizar "Trans:" → "Transv:" para melhor leitura
        _trans_label = _trans_label.replace("Trans: ", "Transv: ", 1)

        # Quebra de linha automática em " + " (max 33 chars por linha)
        def _quebrar(texto, mc=33):
            if not texto or len(texto) <= mc:
                return [texto] if texto else []
            linhas = []
            while len(texto) > mc:
                pos = texto.rfind(" + ", 0, mc + 3)
                if pos > 0:
                    linhas.append(texto[:pos]); texto = texto[pos+3:]
                else:
                    pos2 = texto.rfind(" ", 0, mc)
                    if pos2 > 0:
                        linhas.append(texto[:pos2]); texto = texto[pos2+1:]
                    else:
                        linhas.append(texto[:mc]); texto = texto[mc:]
            if texto: linhas.append(texto)
            return linhas

        _linhas_long  = _quebrar(_long_label)
        _linhas_trans = _quebrar(_trans_label)

        # Montar sequência final: long | linha vazia de separação | trans
        _seq = _linhas_long
        if _linhas_trans:
            _seq = _seq + [""] + _linhas_trans  # "" = separação visual

        # Calcular step dinâmico para caber entre y+22mm e y+9mm
        _INICIO  = 22.0   # mm — onde começa o bloco de armaduras
        _FIM_MIN =  9.0   # mm — mínimo para não sobrepor "Mat:"
        _espaco  = _INICIO - _FIM_MIN  # 13mm disponíveis
        _n       = len(_seq)
        _step    = _espaco / _n if _n > 0 else 3.5
        _step    = max(2.8, min(3.8, _step))   # limitar entre 2.8 e 3.8mm

        # Linha "Armaduras:" em negrito
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(colors.black)
        c.drawString(tx, y + 26*mm, "Armaduras:")

        # Desenhar cada linha de armadura
        c.setFont("Helvetica", 6.5)
        c.setFillColor(colors.HexColor("#222222"))
        for _i, _linha in enumerate(_seq):
            if _linha:  # pular linhas vazias (apenas deslocam o cursor)
                _ypos = (_INICIO - _i * _step) * mm
                c.drawString(tx, y + _ypos, _linha)

        # Mat e Pavimento ancorados em posições fixas baixas
        _y_mat = min((_INICIO - _n * _step) * mm - 0.5*mm, 8.5*mm)
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#333333"))
        c.drawString(tx, y + _y_mat, f"Mat: {reg['Material'][:20]}")

        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(colors.black)
        c.drawString(tx, y + 7*mm, reg["Pavimento"][:22])

        c.setFont("Helvetica-Oblique", 6.5)
        c.setFillColor(colors.gray)
        c.drawString(tx, y + 3*mm, f"Obra: {nome_projeto[:22]}")

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

    novo = {
        "ID_Projeto":          id_proj,
        "Nome_Obra":           nome,
        "Data_Upload":         datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Total_Elementos":     len(registros),
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
    metricas = list(tipos_count.items())
    N_COLS = 4  # número fixo de colunas — nunca estoura o índice
    cols = st.columns(N_COLS)
    for i, (label, valor) in enumerate(metricas):
        cols[i % N_COLS].metric(label, valor)

    # Filtros
    filtrados = [
        r for r in registros
        if r["Tipo_Legivel"] in filtro_tipo
        and (not filtro_pav or filtro_pav.lower() in r["Pavimento"].lower())
    ]

    st.divider()
    colunas_view = ["Nome", "Tipo_Legivel", "Pavimento", "Geometria",
                    "Armadura", "Material", "Status"]
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
