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

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- CACHE GLOBAL PARA OTIMIZAR ---
CACHE_ARMADURAS = {}

# --- FUNÇÕES DE CONEXÃO ---
def conectar_google_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists(ARQUIVO_CREDENCIAIS):
        creds = Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
    else:
        st.error("Credenciais não encontradas.")
        st.stop()
    return gspread.authorize(creds)

def limpar_string(texto):
    if not texto: return "X"
    return "".join(e for e in str(texto) if e.isalnum()).upper()

# --- NOVA LÓGICA DE ARMADURA (GLOBAL MATCH) ---
def pre_carregar_armaduras(ifc_file):
    """Lê todas as barras do arquivo DE UMA VEZ e organiza por dono (P1, P2...)."""
    global CACHE_ARMADURAS
    CACHE_ARMADURAS = {}
    
    # Pega todas as barras do projeto inteiro
    barras = ifc_file.by_type('IfcReinforcingBar')
    
    for bar in barras:
        nome_completo = bar.Name # Ex: "1 P1 \X\D810.00 C=280.00"
        if not nome_completo: continue
        
        # Tenta achar o nome do elemento no texto da barra
        # Procura padrão: Número + Espaço + Pxxx + Espaço
        # Ex: "1 P4 " -> Pega "P4"
        match_nome = re.search(r'^\d+\s+(P\d+)\s+', nome_completo)
        
        if match_nome:
            nome_elemento = match_nome.group(1) # "P4"
            
            # Extrai Bitola
            match_bitola = re.search(r'\\X\\D8\s*([0-9\.]+)', nome_completo)
            bitola = 0.0
            if match_bitola:
                bitola = float(match_bitola.group(1))
            else:
                # Tenta propriedade física
                if hasattr(bar, "NominalDiameter") and bar.NominalDiameter:
                    bitola = bar.NominalDiameter * 1000
            
            if bitola > 0:
                if nome_elemento not in CACHE_ARMADURAS:
                    CACHE_ARMADURAS[nome_elemento] = []
                CACHE_ARMADURAS[nome_elemento].append(bitola)

def get_armadura_pilar(nome_pilar):
    """Busca a armadura no cache pelo nome do pilar (ex: P4)."""
    # Tenta match exato
    lista_bitolas = CACHE_ARMADURAS.get(nome_pilar)
    
    if not lista_bitolas:
        return "Não detalhado / Sem barras 3D"
        
    from collections import Counter
    c = Counter(lista_bitolas)
    # Ordena decrescente
    return " + ".join([f"{qtd} ø{diam:.1f}" for diam, qtd in sorted(c.items(), reverse=True)])

# --- NOVA LÓGICA DE GEOMETRIA (BOUNDING BOX) ---
def extrair_secao_universal(pilar):
    """Mede a geometria 3D ponto a ponto (funciona para qualquer formato)."""
    # 1. Tenta Psets TQS (Mais rápido/preciso se existir)
    psets = ifcopenshell.util.element.get_psets(pilar)
    if 'TQS_Geometria' in psets:
        d = psets['TQS_Geometria']
        b = d.get('Dimensao_b1') or d.get('B')
        h = d.get('Dimensao_h1') or d.get('H')
        if b and h:
            vals = sorted([float(b), float(h)])
            if vals[0] < 2.0: vals = [v*100 for v in vals] # Converte m para cm
            return f"{vals[0]:.0f}x{vals[1]:.0f}"

    # 2. Força Bruta: Bounding Box 3D
    # Coleta todos os pontos cartesianos que formam o pilar
    if not pilar.Representation: return "N/A"
    
    pontos_x = []
    pontos_y = []
    
    # Função recursiva para achar pontos escondidos em malhas
    def buscar_pontos(item):
        if item.is_a('IfcCartesianPoint'):
            if len(item.Coordinates) >= 2:
                pontos_x.append(item.Coordinates[0])
                pontos_y.append(item.Coordinates[1])
        # Navega para dentro das estruturas
        if hasattr(item, 'Points'): 
            for p in item.Points: buscar_pontos(p)
        if hasattr(item, 'OuterCurve'): buscar_pontos(item.OuterCurve)
        if hasattr(item, 'PolygonalBoundary'): buscar_pontos(item.PolygonalBoundary)
        if hasattr(item, 'FbsmFaces'): # Malha TQS
            for face in item.FbsmFaces: buscar_pontos(face)
        if hasattr(item, 'CfsFaces'): 
            for face in item.CfsFaces: buscar_pontos(face)
        if hasattr(item, 'Bounds'): 
            for b in item.Bounds: buscar_pontos(b)
            
    for rep in pilar.Representation.Representations:
        if rep.RepresentationIdentifier in ['Body', 'Mesh']:
            for item in rep.Items:
                buscar_pontos(item)
    
    if pontos_x and pontos_y:
        largura = max(pontos_x) - min(pontos_x)
        altura = max(pontos_y) - min(pontos_y)
        
        # Ajuste de unidade (metros para cm)
        if largura < 2.0: largura *= 100
        if altura < 2.0: altura *= 100
        
        # Arredonda e ordena
        dims = sorted([largura, altura])
        # Tolerância para arredondamento (ex: 13.99 vira 14)
        return f"{dims[0]:.0f}x{dims[1]:.0f}"
        
    return "N/A"

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    
    # 1. INDEXAÇÃO GLOBAL DAS ARMADURAS (O Segredo!)
    pre_carregar_armaduras(ifc_file)
    
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

        # Usar as novas funções "Universais"
        secao = extrair_secao_universal(pilar)
        
        # Pega armadura buscando pelo NOME (P1, P2...)
        armadura = get_armadura_pilar(nome)

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
    largura_pag, altura_pag = A4
    largura_etq, altura_etq = 90*mm, 50*mm
    margem, espaco = 10*mm, 5*mm
    x, y = margem, altura_pag - margem - altura_etq
    
    for pilar in dados_pilares:
        c.setLineWidth(0.5)
        c.rect(x, y, largura_etq, altura_etq)
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        temp_qr_path = f"temp_{pilar['ID_Unico'][:5]}.png"
        img_qr.save(temp_qr_path)
        c.drawImage(temp_qr_path, x+2*mm, y+5*mm, width=40*mm, height=40*mm)
        os.remove(temp_qr_path)
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x+45*mm, y+20*mm, f"Pav: {pilar['Pavimento']}")
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(x+45*mm, y+10*mm, f"Proj: {nome_projeto_legivel[:15]}")
        
        x += largura_etq + espaco
        if x + largura_etq > largura_pag - margem:
            x = margem
            y -= (altura_etq + espaco)
        if y < margem:
            c.showPage()
            x = margem
            y = altura_pag - margem - altura_etq
    c.save()
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(page_title="Gestor BIM", page_icon="🏗️")
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        s = st.text_input("Senha", type="password")
        if st.button("Entrar"):
            if s == "bim123":
                st.session_state['logado'] = True
                st.rerun()
            else: st.error("Senha incorreta")
        return

    st.title("🏗️ Gestor BIM (Relacional)")
    nome_projeto_legivel = st.text_input("Nome da Obra", placeholder="Ex: Edifício Diogenes")
    id_projeto = limpar_string(nome_projeto_legivel)
    arquivo_upload = st.file_uploader("Carregar IFC", type=["ifc"])
    
    if arquivo_upload and nome_projeto_legivel:
        if st.button("🚀 PROCESSAR", type="primary"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                with st.spinner('Lendo IFC e Armaduras...'):
                    dados_pilares = processar_ifc(caminho_temp, id_projeto)
                os.remove(caminho_temp)

                with st.spinner('Atualizando Banco de Dados...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    
                    try: ws_proj = sh.worksheet("Projetos")
                    except: ws_proj = sh.add_worksheet("Projetos", 100, 5)
                    lista_proj = ws_proj.get_all_records()
                    df_proj = pd.DataFrame(lista_proj)
                    if not df_proj.empty and 'ID_Projeto' in df_proj.columns:
                        df_proj = df_proj[df_proj['ID_Projeto'] != id_projeto]
                    novo_proj = {'ID_Projeto': id_projeto, 'Nome_Obra': nome_projeto_legivel, 'Data_Upload': datetime.datetime.now().strftime("%d/%m/%Y"), 'Total_Pilares': len(dados_pilares)}
                    df_proj_final = pd.concat([df_proj, pd.DataFrame([novo_proj])], ignore_index=True)
                    ws_proj.clear()
                    ws_proj.update([df_proj_final.columns.values.tolist()] + df_proj_final.values.tolist())

                    try: ws_pil = sh.worksheet("Pilares")
                    except: ws_pil = sh.add_worksheet("Pilares", 1000, 10)
                    lista_pil = ws_pil.get_all_records()
                    df_pil = pd.DataFrame(lista_pil)
                    if not df_pil.empty and 'Projeto_Ref' in df_pil.columns:
                        df_pil = df_pil[df_pil['Projeto_Ref'] != id_projeto]
                    df_pil_novos = pd.DataFrame(dados_pilares)
                    df_pil_final = pd.concat([df_pil, df_pil_novos], ignore_index=True)
                    ws_pil.clear()
                    ws_pil.update([df_pil_final.columns.values.tolist()] + df_pil_final.values.tolist())

                with st.spinner('Gerando PDF...'):
                    pdf_buffer = gerar_pdf_memoria(dados_pilares, nome_projeto_legivel)
                
                st.success(f"✅ Projeto '{nome_projeto_legivel}' processado! {len(dados_pilares)} pilares encontrados.")
                st.download_button("📥 BAIXAR PDF", pdf_buffer, f"Etiquetas_{id_projeto}.pdf", "application/pdf")
                
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
