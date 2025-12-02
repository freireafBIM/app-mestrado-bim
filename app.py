import streamlit as st
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
import io
import toml

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"
NOME_PASTA_DRIVE = "Etiquetas_BIM_Projetos" # Nome da pasta que será criada no Drive

# --- FUNÇÕES DE CONEXÃO ---

def obter_credenciais():
    """Retorna o objeto de credenciais (Local ou Nuvem)."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        return Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists(ARQUIVO_CREDENCIAIS):
        return Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
    else:
        st.error("Credenciais não encontradas!")
        return None

def conectar_google_sheets():
    creds = obter_credenciais()
    client = gspread.authorize(creds)
    return client

def enviar_pdf_drive(pdf_buffer, nome_arquivo):
    """Envia o PDF para o Google Drive e retorna o Link Público."""
    creds = obter_credenciais()
    service = build('drive', 'v3', credentials=creds)
    
    # 1. Verifica/Cria a pasta no Drive para organizar
    query = f"name='{NOME_PASTA_DRIVE}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    items = results.get('files', [])
    
    if not items:
        # Cria a pasta se não existir
        file_metadata = {
            'name': NOME_PASTA_DRIVE,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        folder_id = folder.get('id')
    else:
        folder_id = items[0]['id']

    # 2. Faz o Upload do Arquivo
    file_metadata = {
        'name': nome_arquivo,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(pdf_buffer, mimetype='application/pdf', resumable=True)
    
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    file_id = file.get('id')
    web_link = file.get('webViewLink') # Link para abrir no navegador
    
    # 3. Permissões (Deixar público para quem tem o link ler)
    # Isso é crucial para o AppSheet conseguir abrir
    service.permissions().create(
        fileId=file_id,
        body={'role': 'reader', 'type': 'anyone'}
    ).execute()
    
    return web_link

# --- FUNÇÕES DE LÓGICA DE NEGÓCIO ---

def extrair_texto_armadura(pilar):
    """Inferência de armadura."""
    barras = []
    relacoes = getattr(pilar, 'IsDecomposedBy', [])
    for rel in relacoes:
        if rel.is_a('IfcRelAggregates'):
            for obj in rel.RelatedObjects:
                if obj.is_a('IfcReinforcingBar'):
                    d = round(obj.NominalDiameter * 1000, 1)
                    barras.append(d)
    
    if not barras:
        psets = ifcopenshell.util.element.get_psets(pilar)
        for nome, dados in psets.items():
            if 'Armadura' in nome or 'Reinforcement' in nome:
                for k, v in dados.items():
                    if isinstance(v, str) and len(v) > 5: return v
        return "Verificar Projeto (Sem vínculo 3D)"
    
    from collections import Counter
    c = Counter(barras)
    return " + ".join([f"{qtd} ø{diam}" for diam, qtd in c.items()])

def processar_ifc(caminho_arquivo, nome_projeto_input):
    """Processa IFC e retorna dados."""
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
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
        
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name

        dados.append({
            'Projeto': nome_projeto_input, 
            'ID_Unico': guid, 
            'Nome': nome, 
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': '',
            'Link_PDF': '' # Placeholder, será preenchido depois
        })
    
    dados.sort(key=lambda x: x['Nome'])
    return dados

def gerar_pdf_memoria(dados_pilares, nome_projeto):
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
        temp_qr_path = f"temp_{pilar['ID_Unico'][:4]}.png"
        img_qr.save(temp_qr_path)
        
        c.drawImage(temp_qr_path, x+2*mm, y+5*mm, width=40*mm, height=40*mm)
        os.remove(temp_qr_path)
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        c.drawString(x+45*mm, y+20*mm, f"Pav: {pilar['Pavimento']}")
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(x+45*mm, y+10*mm, f"Obra: {nome_projeto[:15]}")
        
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
    st.set_page_config(page_title="Gestor Multi-Obras BIM", page_icon="🏗️")
    
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

    st.title("🏗️ Gestor Multi-Obras BIM")
    if st.sidebar.button("Sair"):
        st.session_state['logado'] = False
        st.rerun()

    nome_projeto = st.text_input("Nome do Projeto / Obra", placeholder="Ex: Ed. Diogenes e Kely")
    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload is not None and nome_projeto:
        if st.button("🚀 PROCESSAR, GERAR PDF E SALVAR", type="primary"):
            try:
                # 1. Processar IFC
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                with st.spinner('Extraindo dados do BIM...'):
                    novos_dados = processar_ifc(caminho_temp, nome_projeto)
                os.remove(caminho_temp)

                # 2. Gerar PDF na memória
                with st.spinner('Gerando Etiquetas PDF...'):
                    pdf_buffer = gerar_pdf_memoria(novos_dados, nome_projeto)

                # 3. Enviar PDF para Google Drive
                with st.spinner('Enviando PDF para Google Drive...'):
                    nome_arquivo_pdf = f"Etiquetas_{nome_projeto}.pdf"
                    link_publico = enviar_pdf_drive(pdf_buffer, nome_arquivo_pdf)
                
                # 4. Atualizar os dados com o Link
                for item in novos_dados:
                    item['Link_PDF'] = link_publico

                # 5. Salvar na Planilha
                with st.spinner('Atualizando Banco de Dados...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    ws = sh.sheet1
                    
                    dados_existentes = ws.get_all_records()
                    df_antigo = pd.DataFrame(dados_existentes)
                    
                    if not df_antigo.empty and 'Projeto' in df_antigo.columns:
                        df_limpo = df_antigo[df_antigo['Projeto'] != nome_projeto]
                    else:
                        df_limpo = pd.DataFrame()

                    df_novo = pd.DataFrame(novos_dados)
                    df_final = pd.concat([df_limpo, df_novo], ignore_index=True)
                    
                    ws.clear()
                    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())
                
                st.success(f"✅ Sucesso! PDF salvo no Drive e vinculado ao projeto '{nome_projeto}'.")
                st.markdown(f"**[Clique aqui para acessar o PDF gerado]({link_publico})**")
                
                # Reseta o ponteiro do buffer para permitir download direto também
                pdf_buffer.seek(0)
                st.download_button("📥 BAIXAR ETIQUETAS AGORA", pdf_buffer, nome_arquivo_pdf, "application/pdf")
                
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
