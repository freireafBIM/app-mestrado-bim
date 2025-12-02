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
import toml
import re # Biblioteca para limpar texto (Regex)

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
    """Remove espaços e caracteres especiais para criar IDs limpos."""
    if not texto: return "X"
    # Mantém apenas letras e números e deixa maiúsculo
    return "".join(e for e in str(texto) if e.isalnum()).upper()

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

def processar_ifc(caminho_arquivo, nome_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    # Cria sufixo limpo do projeto (Ex: "Ed. Diogenes" -> "EDDIOGENES")
    sufixo_projeto = limpar_string(nome_projeto_input)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        # 1. Pavimento (Extraído ANTES para compor o ID)
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name
        
        # Cria sufixo limpo do pavimento (Ex: "1º Pavimento" -> "1PAVIMENTO")
        sufixo_pav = limpar_string(pavimento)

        # --- NOVA LÓGICA DE CHAVE PRIMÁRIA ---
        # GUID + PAVIMENTO + PROJETO
        # Ex: 3X64...-TERREO-EDDIOGENES
        id_composto = f"{guid}-{sufixo_pav}-{sufixo_projeto}"

        # 2. Geometria
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
            'Projeto': nome_projeto_input, 
            'ID_Unico': id_composto, # <--- AQUI ESTÁ A CHAVE NOVA
            'Nome': nome, 
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    # Ordena por Pavimento e depois por Nome
    dados.sort(key=lambda x: (x['Pavimento'], x['Nome']))
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
        
        # QR Code
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        temp_qr_path = f"temp_{pilar['ID_Unico'][:4]}.png"
        img_qr.save(temp_qr_path)
        
        c.drawImage(temp_qr_path, x+2*mm, y+5*mm, width=40*mm, height=40*mm)
        os.remove(temp_qr_path)
        
        # Textos
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        # Pavimento em destaque
        c.setFont("Helvetica-Bold", 10)
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
    st.set_page_config(page_title="Gestão de Armaduras", page_icon="🏗️")
    
    # LOGIN
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    
    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        senha = st.text_input("Senha de Acesso", type="password")
        if st.button("Entrar"):
            if senha == "bim123":
                st.session_state['logado'] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
        return

    st.title("🏗️ Gestor de Etiquetas BIM")
    
    with st.sidebar:
        st.write("Status: Conectado")
        if st.button("Sair"):
            st.session_state['logado'] = False
            st.rerun()

    nome_projeto = st.text_input("Nome do Projeto / Obra", placeholder="Ex: Ed. Thiago")
    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload and nome_projeto:
        if st.button("🚀 PROCESSAR DADOS", type="primary"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                # A. Processamento
                with st.spinner('Lendo IFC e gerando Chaves Únicas por Pavimento...'):
                    novos_dados = processar_ifc(caminho_temp, nome_projeto)
                os.remove(caminho_temp)

                # B. Google Sheets
                with st.spinner('Sincronizando Banco de Dados...'):
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

                # C. Gerar PDF
                with st.spinner('Gerando etiquetas...'):
                    pdf_buffer = gerar_pdf_memoria(novos_dados, nome_projeto)
                
                st.success(f"✅ Sucesso! {len(novos_dados)} pilares identificados e diferenciados por pavimento.")
                
                # D. Botão de Download
                nome_arquivo_pdf = f"Etiquetas_{nome_projeto}.pdf"
                st.download_button(
                    label="📥 BAIXAR PDF DAS ETIQUETAS",
                    data=pdf_buffer,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf"
                )
                
            except Exception as e:
                st.error(f"Erro no processamento: {e}")

if __name__ == "__main__":
    main()

