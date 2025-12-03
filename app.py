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
import io
import datetime
import re
from collections import Counter

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- CACHE GLOBAL ---
CACHE_ARMADURAS_POR_NOME = {}

# --- FUNÇÕES DE CONEXÃO ---
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

# --- LÓGICA DE EXTRAÇÃO BLINDADA (CORRIGIDA) ---

def indexar_todas_armaduras(ifc_file):
    """Lê todas as barras e extrai dados independente da formatação do texto."""
    global CACHE_ARMADURAS_POR_NOME
    CACHE_ARMADURAS_POR_NOME = {}
    
    barras = ifc_file.by_type('IfcReinforcingBar')
    
    for bar in barras:
        nome_completo = bar.Name # Ex: "1 P1 \X\D810.00" OU "1 P1 Ø10.00"
        if not nome_completo: continue
        
        # 1. Achar o nome do Pilar (P1, P12, P100...)
        # Procura "P" seguido de números em qualquer lugar da string
        match_nome = re.search(r'(P\d+)', nome_completo)
        
        if match_nome:
            nome_pilar = match_nome.group(1) # "P1"
            bitola = 0.0
            
            # 2. Tenta achar a Bitola no Texto (Várias estratégias)
            
            # Estratégia A: Procura o código IFC bruto (\X\D8)
            match_bruto = re.search(r'\\X\\D8\s*([0-9\.]+)', nome_completo)
            
            # Estratégia B: Procura o símbolo Ø ou qualquer caractere estranho seguido de número
            match_simbolo = re.search(r'[Øø]\s*([0-9\.]+)', nome_completo)
            
            if match_bruto:
                bitola = float(match_bruto.group(1))
            elif match_simbolo:
                bitola = float(match_simbolo.group(1))
            
            # Estratégia C (Garantia): Se não achou no texto, pega do objeto físico
            if bitola == 0.0 and hasattr(bar, "NominalDiameter") and bar.NominalDiameter:
                # O TQS exporta NominalDiameter em metros (ex: 0.01)
                bitola = bar.NominalDiameter * 1000
            
            # Se achou algo válido, salva no cache
            if bitola > 0:
                if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
                    CACHE_ARMADURAS_POR_NOME[nome_pilar] = []
                CACHE_ARMADURAS_POR_NOME[nome_pilar].append(bitola)

def obter_armadura_do_cache(nome_pilar):
    """Busca a armadura acumulada para aquele pilar."""
    # Se não achar exato, tenta achar contido (ex: pilar chama "P1", mas no cache tá "P1 (id...)")
    if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
        return "Verificar Detalhamento (Sem barras identificadas)"
    
    lista_bitolas = CACHE_ARMADURAS_POR_NOME[nome_pilar]
    c = Counter(lista_bitolas)
    
    # Formata: "4 ø10.0 + 12 ø5.0"
    return " + ".join([f"{qtd} ø{diam:.1f}" for diam, qtd in sorted(c.items(), reverse=True)])

def extrair_secao_universal(pilar):
    """Mede a geometria 3D ponto a ponto (Bounding Box) - VERSÃO CORRIGIDA."""
    # 1. Tenta Psets TQS primeiro
    psets = ifcopenshell.util.element.get_psets(pilar)
    if 'TQS_Geometria' in psets:
        d = psets['TQS_Geometria']
        b = d.get('Dimensao_b1') or d.get('B')
        h = d.get('Dimensao_h1') or d.get('H')
        if b and h:
            vals = sorted([float(b), float(h)])
            if vals[0] < 3.0: vals = [v*100 for v in vals]
            return f"{vals[0]:.0f}x{vals[1]:.0f}"

    # 2. Varredura 3D (Fallback)
    if not pilar.Representation: return "N/A"
    
    pontos_x, pontos_y = [], []
    
    def coletar_pontos(item):
        # Proteção contra recursão infinita ou tipos inválidos
        if item is None: return
        
        # Se for lista ou tupla, navega dentro
        if isinstance(item, (list, tuple)):
            for i in item: coletar_pontos(i)
            return

        # Verifica se é entidade IFC válida antes de chamar .is_a()
        if not hasattr(item, 'is_a'): return

        # Se for Ponto Cartesiano
        if item.is_a('IfcCartesianPoint'):
            if hasattr(item, 'Coordinates') and len(item.Coordinates) >= 2:
                pontos_x.append(item.Coordinates[0])
                pontos_y.append(item.Coordinates[1])
            return

        # Lista de atributos para explorar recursivamente
        atributos_para_explorar = [
            'Points', 'OuterCurve', 'PolygonalBoundary', 
            'FbsmFaces', 'CfsFaces', 'Bounds', 'Bound', 
            'Items', 'MappingSource', 'MappedRepresentation', 
            'Polygon', 'SweptArea'
        ]
        
        for attr in atributos_para_explorar:
            if hasattr(item, attr):
                val = getattr(item, attr)
                coletar_pontos(val)

    # Inicia a varredura
    for rep in pilar.Representation.Representations:
        if rep.RepresentationIdentifier in ['Body', 'Mesh', 'Box', 'Facetation']:
            for item in rep.Items:
                coletar_pontos(item)
    
    # Calcula dimensão final
    if pontos_x and pontos_y:
        try:
            largura = max(pontos_x) - min(pontos_x)
            altura = max(pontos_y) - min(pontos_y)
            
            # Filtro de sanidade (evitar 0x0)
            if largura <= 0 or altura <= 0: return "N/A"

            if largura < 3.0: largura *= 100
            if altura < 3.0: altura *= 100
            
            dims = sorted([largura, altura])
            return f"{dims[0]:.0f}x{dims[1]:.0f}"
        except:
            return "N/A"

    return "N/A"
    
    def coletar_pontos(item):
        if item.is_a('IfcCartesianPoint') and len(item.Coordinates) >= 2:
            pontos_x.append(item.Coordinates[0])
            pontos_y.append(item.Coordinates[1])
        # Recursividade para entrar em listas e sub-objetos
        atributos_lista = ['Points', 'OuterCurve', 'PolygonalBoundary', 'FbsmFaces', 'CfsFaces', 'Bounds', 'Items', 'MappingSource', 'MappedRepresentation']
        for attr in atributos_lista:
            if hasattr(item, attr):
                val = getattr(item, attr)
                if isinstance(val, list):
                    for v in val: coletar_pontos(v)
                else:
                    coletar_pontos(val)
        if hasattr(item, 'Bound'): coletar_pontos(item.Bound)
        if hasattr(item, 'Polygon'): 
            for pt in item.Polygon: coletar_pontos(pt)

    for rep in pilar.Representation.Representations:
        if rep.RepresentationIdentifier in ['Body', 'Mesh', 'Box']:
            for item in rep.Items:
                coletar_pontos(item)
    
    if pontos_x and pontos_y:
        largura = max(pontos_x) - min(pontos_x)
        altura = max(pontos_y) - min(pontos_y)
        if largura < 3.0: largura *= 100
        if altura < 3.0: altura *= 100
        dims = sorted([largura, altura])
        return f"{dims[0]:.0f}x{dims[1]:.0f}"
        
    return "N/A"

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    
    # 1. INDEXAR TODAS AS BARRAS ANTES
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

        secao = extrair_secao_universal(pilar)
        armadura = obter_armadura_do_cache(nome)

        dados.append({
            'ID_Unico': id_unico_pilar,   
            'Projeto_Ref': id_projeto_input, 
            'Nome': nome, 
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    dados.sort(key=lambda x: (x['Pavimento'], x['Nome']))
    return dados

def gerar_pdf_memoria(dados_pilares, nome_projeto_legivel):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    x, y = 10*mm, 297*mm - 10*mm - 50*mm
    
    for pilar in dados_pilares:
        c.setLineWidth(0.5)
        c.rect(x, y, 90*mm, 50*mm)
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save("temp_qr.png")
        c.drawImage("temp_qr.png", x+2*mm, y+5*mm, width=40*mm, height=40*mm)
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        c.drawString(x+45*mm, y+20*mm, f"Pav: {pilar['Pavimento']}")
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(x+45*mm, y+10*mm, f"Proj: {nome_projeto_legivel[:15]}")
        
        x += 95*mm
        if x > 210*mm - 10*mm:
            x = 10*mm
            y -= 55*mm
        if y < 10*mm:
            c.showPage()
            x, y = 10*mm, 297*mm - 60*mm
            
    c.save()
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(page_title="Gestor BIM", page_icon="🏗️")
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    
    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        if st.text_input("Senha", type="password") == "bim123" and st.button("Entrar"):
            st.session_state['logado'] = True
            st.rerun()
        return

    st.title("🏗️ Gestor BIM (Relacional)")
    nome = st.text_input("Nome da Obra")
    id_proj = limpar_string(nome)
    f = st.file_uploader("IFC", type=["ifc"])
    
    if f and nome:
        if st.button("🚀 PROCESSAR"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as t:
                t.write(f.getvalue())
                path = t.name
            
            with st.spinner('Processando...'):
                dados = processar_ifc(path, id_proj)
            os.remove(path)
            
            client = conectar_google_sheets()
            sh = client.open(NOME_PLANILHA_GOOGLE)
            
            # Atualiza PROJETOS
            try: ws_p = sh.worksheet("Projetos")
            except: ws_p = sh.add_worksheet("Projetos", 100, 5)
            recs = ws_p.get_all_records()
            df = pd.DataFrame(recs)
            if not df.empty and 'ID_Projeto' in df.columns: df = df[df['ID_Projeto'] != id_proj]
            new = {'ID_Projeto': id_proj, 'Nome_Obra': nome, 'Data_Upload': str(datetime.date.today()), 'Total_Pilares': len(dados)}
            df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
            ws_p.clear()
            ws_p.update([df.columns.values.tolist()] + df.values.tolist())

            # Atualiza PILARES
            try: ws_pil = sh.worksheet("Pilares")
            except: ws_pil = sh.add_worksheet("Pilares", 1000, 10)
            recs = ws_pil.get_all_records()
            df = pd.DataFrame(recs)
            if not df.empty and 'Projeto_Ref' in df.columns: df = df[df['Projeto_Ref'] != id_proj]
            df = pd.concat([df, pd.DataFrame(dados)], ignore_index=True)
            ws_pil.clear()
            ws_pil.update([df.columns.values.tolist()] + df.values.tolist())
            
            pdf = gerar_pdf_memoria(dados, nome)
            st.success("Sucesso!")
            st.download_button("Baixar PDF", pdf, "etiquetas.pdf", "application/pdf")

if __name__ == "__main__":
    main()

