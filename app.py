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
import io # Para lidar com arquivos na memória

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- FUNÇÕES DE BACKEND (Lógica do Python) ---

import toml # Adicione no topo se precisar, mas st.secrets já resolve

def conectar_google_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Tenta usar o cofre da nuvem (Streamlit Secrets)
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    
    # Se não achar, tenta usar o arquivo local (Seu PC)
    elif os.path.exists(ARQUIVO_CREDENCIAIS):
        creds = Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
    
    else:
        st.error("Arquivo de credenciais não encontrado!")
        return None

    client = gspread.authorize(creds)
    return client

def extrair_texto_armadura(pilar):
    """Lógica de inferência de armadura."""
    barras_encontradas = []
    relacoes = getattr(pilar, 'IsDecomposedBy', [])
    for rel in relacoes:
        if rel.is_a('IfcRelAggregates'):
            for obj in rel.RelatedObjects:
                if obj.is_a('IfcReinforcingBar'):
                    diam = round(obj.NominalDiameter * 1000, 1)
                    barras_encontradas.append(diam)
    
    if not barras_encontradas:
        # Tenta ler Psets de texto como fallback
        psets = ifcopenshell.util.element.get_psets(pilar)
        for nome, dados in psets.items():
            if 'Armadura' in nome or 'Reinforcement' in nome:
                for k, v in dados.items():
                    if isinstance(v, str) and len(v) > 5: return v
        return "Verificar Projeto (Sem vínculo 3D)"
    
    from collections import Counter
    c = Counter(barras_encontradas)
    return " + ".join([f"{qtd} ø{diam}" for diam, qtd in c.items()])

def processar_ifc(caminho_arquivo):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0) # Barra de progresso visual
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        # Atualiza a barra de progresso
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        # Extração de Seção (Lógica Geométrica)
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
            'ID_Unico': guid, 'Nome': nome, 'Secao': secao,
            'Armadura': armadura, 'Pavimento': pavimento,
            'Status': 'A CONFERIR', 'Data_Conferencia': '', 'Responsavel': ''
        })
    
    dados.sort(key=lambda x: x['Nome'])
    return dados

def gerar_pdf_memoria(dados_pilares):
    """Gera o PDF diretamente na memória RAM (sem salvar no disco)."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largura_pag, altura_pag = A4
    largura_etq, altura_etq = 90*mm, 50*mm
    margem, espaco = 10*mm, 5*mm
    
    x, y = margem, altura_pag - margem - altura_etq
    
    for pilar in dados_pilares:
        c.setLineWidth(0.5)
        c.rect(x, y, largura_etq, altura_etq)
        
        # QR Code temporário
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        temp_qr_path = f"temp_{pilar['ID_Unico'][:4]}.png"
        img_qr.save(temp_qr_path)
        
        c.drawImage(temp_qr_path, x+2*mm, y+5*mm, width=40*mm, height=40*mm)
        os.remove(temp_qr_path) # Limpa o lixo
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        c.drawString(x+45*mm, y+20*mm, f"Pav: {pilar['Pavimento']}")
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(x+45*mm, y+10*mm, "Projeto Mestrado BIM")
        
        x += largura_etq + espaco
        if x + largura_etq > largura_pag - margem:
            x = margem
            y -= (altura_etq + espaco)
        if y < margem:
            c.showPage()
            x = margem
            y = altura_pag - margem - altura_etq
            
    c.save()
    buffer.seek(0) # Retorna o ponteiro para o início do arquivo na memória
    return buffer

# --- FRONTEND (INTERFACE WEB) ---

def main():
    st.set_page_config(page_title="Gestor de Etiquetas BIM", page_icon="🏗️")
    
    # 1. TELA DE LOGIN SIMPLES
    if 'logado' not in st.session_state:
        st.session_state['logado'] = False

    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        senha = st.text_input("Digite a senha de acesso:", type="password")
        if st.button("Entrar"):
            if senha == "bim123": # <--- SUA SENHA AQUI
                st.session_state['logado'] = True
                st.rerun() # Recarrega a página
            else:
                st.error("Senha incorreta.")
        return # Para a execução aqui se não estiver logado

    # 2. TELA PRINCIPAL (SÓ APARECE SE LOGADO)
    st.title("🏗️ Gerador de Etiquetas & Integração BIM")
    st.markdown("Faça o upload do arquivo IFC gerado no TQS para sincronizar com o Google Sheets e gerar etiquetas.")
    
    # Upload de Arquivo
    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload is not None:
        st.info(f"Arquivo carregado: {arquivo_upload.name}")
        
        if st.button("🚀 PROCESSAR ARQUIVO", type="primary"):
            try:
                # Salva o arquivo uploadado temporariamente no disco para o ifcopenshell ler
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                with st.spinner('Minerando dados do IFC...'):
                    dados = processar_ifc(caminho_temp)
                
                with st.spinner('Enviando para Google Sheets...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    ws = sh.sheet1
                    df = pd.DataFrame(dados)
                    ws.clear()
                    ws.update([df.columns.values.tolist()] + df.values.tolist())
                
                with st.spinner('Gerando PDF...'):
                    pdf_buffer = gerar_pdf_memoria(dados)
                
                # Sucesso
                st.success(f"✅ Sucesso! {len(dados)} pilares processados e planilha atualizada.")
                
                # Botão de Download
                st.download_button(
                    label="📥 BAIXAR ETIQUETAS (PDF)",
                    data=pdf_buffer,
                    file_name="Etiquetas_Obra.pdf",
                    mime="application/pdf"
                )
                
                # Limpeza
                os.remove(caminho_temp)
                
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    main()