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

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

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
        return None

    client = gspread.authorize(creds)
    return client

# --- LÓGICA DE EXTRAÇÃO BIM ---
def limpar_string(texto):
    if not texto: return "X"
    return "".join(e for e in str(texto) if e.isalnum()).upper()

def extrair_texto_armadura(pilar):
    barras = []
    relacoes = getattr(pilar, 'IsDecomposedBy', [])
    for rel in relacoes:
        if rel.is_a('IfcRelAggregates'):
            for obj in rel.RelatedObjects:
                if obj.is_a('IfcReinforcingBar'):
                    d = round(obj.NominalDiameter * 1000, 1)
                    barras.append(d)
    
    if barras:
        from collections import Counter
        c = Counter(barras)
        return " + ".join([f"{qtd} ø{diam}" for diam, qtd in c.items()])
    
    psets = ifcopenshell.util.element.get_psets(pilar)
    for nome, dados in psets.items():
        if 'Armadura' in nome or 'Reinforcement' in nome:
            for k, v in dados.items():
                if isinstance(v, str) and len(v) > 5: return v
    return "Verificar Projeto (Sem vínculo 3D)"

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        # Pavimento
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name
        
        sufixo_pav = limpar_string(pavimento)

        # CHAVE ÚNICA DO PILAR (ID Composto)
        id_unico_pilar = f"{guid}-{sufixo_pav}-{id_projeto_input}"

        # Geometria
        secao = "N/A"
        if pilar.Representation:
            for rep in pilar.Representation.Representations:
                if rep.RepresentationIdentifier == 'Body':
                    for item in rep.Items:
                        if item.is_a('IfcExtrudedAreaSolid'):
                            perfil = item.SweptArea
                            if perfil.is_a('IfcRectangleProfileDef'):
                                dims = sorted([perfil.XDim * 100, perfil.YDim * 100])
                                secao = f"{dims[0]:.0f}x{dims[1]:.0f}"

        armadura = extrair_texto_armadura(pilar)

        dados.append({
            'ID_Unico': id_unico_pilar,   # Chave Primária da Tabela Pilares
            'Projeto_Ref': id_projeto_input, # Chave Estrangeira (Vínculo com a Tabela Projetos)
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

# --- FRONTEND ---

def main():
    st.set_page_config(page_title="Gestor BIM Relacional", page_icon="🏗️")
    
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
    
    # Inputs
    nome_projeto_legivel = st.text_input("Nome da Obra (Legível)", placeholder="Ex: Edifício Diogenes")
    # Cria um ID técnico para o banco de dados (sem espaços)
    id_projeto = limpar_string(nome_projeto_legivel)
    
    if nome_projeto_legivel:
        st.caption(f"ID Técnico do Projeto: {id_projeto}")

    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload and nome_projeto_legivel:
        if st.button("🚀 PROCESSAR E ATUALIZAR BANCO DE DADOS", type="primary"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                # 1. Processar Pilares
                with st.spinner('Lendo IFC...'):
                    dados_pilares = processar_ifc(caminho_temp, id_projeto)
                os.remove(caminho_temp)

                # 2. Atualizar Google Sheets (2 Tabelas)
                with st.spinner('Sincronizando Tabelas (Projetos e Pilares)...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    
                    # --- ATUALIZA TABELA PROJETOS ---
                    try:
                        ws_proj = sh.worksheet("Projetos")
                    except:
                        ws_proj = sh.add_worksheet("Projetos", 100, 5)

                    # Lê projetos existentes
                    lista_proj = ws_proj.get_all_records()
                    df_proj = pd.DataFrame(lista_proj)
                    
                    # Remove se já existe esse projeto (para atualizar dados)
                    if not df_proj.empty and 'ID_Projeto' in df_proj.columns:
                        df_proj = df_proj[df_proj['ID_Projeto'] != id_projeto]
                    
                    # Cria linha do novo projeto
                    novo_proj = {
                        'ID_Projeto': id_projeto,
                        'Nome_Obra': nome_projeto_legivel,
                        'Data_Upload': datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                        'Total_Pilares': len(dados_pilares)
                    }
                    df_proj_final = pd.concat([df_proj, pd.DataFrame([novo_proj])], ignore_index=True)
                    ws_proj.clear()
                    ws_proj.update([df_proj_final.columns.values.tolist()] + df_proj_final.values.tolist())

                    # --- ATUALIZA TABELA PILARES ---
                    try:
                        ws_pil = sh.worksheet("Pilares")
                    except:
                        ws_pil = sh.add_worksheet("Pilares", 1000, 10)
                    
                    lista_pil = ws_pil.get_all_records()
                    df_pil = pd.DataFrame(lista_pil)
                    
                    # Remove pilares antigos DESTE projeto específico
                    if not df_pil.empty and 'Projeto_Ref' in df_pil.columns:
                        df_pil = df_pil[df_pil['Projeto_Ref'] != id_projeto]
                    
                    # Adiciona novos
                    df_pil_novos = pd.DataFrame(dados_pilares)
                    df_pil_final = pd.concat([df_pil, df_pil_novos], ignore_index=True)
                    
                    ws_pil.clear()
                    ws_pil.update([df_pil_final.columns.values.tolist()] + df_pil_final.values.tolist())

                # 3. PDF
                with st.spinner('Gerando PDF...'):
                    pdf_buffer = gerar_pdf_memoria(dados_pilares, nome_projeto_legivel)
                
                st.success(f"✅ Projeto '{nome_projeto_legivel}' atualizado!")
                st.download_button("📥 BAIXAR PDF", pdf_buffer, f"Etiquetas_{id_projeto}.pdf", "application/pdf")
                
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
