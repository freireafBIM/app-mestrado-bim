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

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- FUNÇÕES DE BACKEND ---

def conectar_google_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Lógica Híbrida: Tenta Nuvem (Secrets) primeiro, depois Local (JSON)
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
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
        psets = ifcopenshell.util.element.get_psets(pilar)
        for nome, dados in psets.items():
            if 'Armadura' in nome or 'Reinforcement' in nome:
                for k, v in dados.items():
                    if isinstance(v, str) and len(v) > 5: return v
        return "Verificar Projeto (Sem vínculo 3D)"
    
    from collections import Counter
    c = Counter(barras_encontradas)
    return " + ".join([f"{qtd} ø{diam}" for diam, qtd in c.items()])

def processar_ifc(caminho_arquivo, nome_projeto_input):
    """
    Processa o IFC e adiciona a coluna PROJETO.
    """
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
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
        
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name

        # ADICIONA O NOME DO PROJETO NO DICIONÁRIO
        dados.append({
            'Projeto': nome_projeto_input, 
            'ID_Unico': guid, 
            'Nome': nome, 
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    dados.sort(key=lambda x: x['Nome'])
    return dados

def gerar_pdf_memoria(dados_pilares, nome_projeto):
    """Gera o PDF com o nome do projeto correto."""
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
        
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x+45*mm, y+35*mm, f"PILAR: {pilar['Nome']}")
        c.setFont("Helvetica", 10)
        c.drawString(x+45*mm, y+25*mm, f"Sec: {pilar['Secao']}")
        c.drawString(x+45*mm, y+20*mm, f"Pav: {pilar['Pavimento']}")
        c.setFont("Helvetica-Oblique", 8)
        # Usa o nome do projeto digitado pelo usuário
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

# --- FRONTEND (INTERFACE WEB) ---

def main():
    st.set_page_config(page_title="Gestor Multi-Obras BIM", page_icon="🏗️")
    
    # --- SISTEMA DE LOGIN COM SENHA (RESTAURADO) ---
    if 'logado' not in st.session_state:
        st.session_state['logado'] = False

    if not st.session_state['logado']:
        st.title("🔒 Acesso Restrito")
        st.markdown("Área exclusiva para gestores BIM.")
        
        senha = st.text_input("Digite a senha de acesso:", type="password")
        
        if st.button("Entrar no Sistema"):
            if senha == "bim123": # <--- SUA SENHA AQUI
                st.session_state['logado'] = True
                st.rerun() # Recarrega a página para entrar
            else:
                st.error("Senha incorreta. Tente novamente.")
        return # Para a execução aqui se não estiver logado

    # --- TELA PRINCIPAL (APÓS LOGIN) ---
    st.title("🏗️ Gestor Multi-Obras BIM")
    st.markdown("Carregue novos projetos para a base de dados central.")
    
    # Botão de Logout (Opcional, mas útil)
    if st.sidebar.button("Sair / Logout"):
        st.session_state['logado'] = False
        st.rerun()

    # 1. INPUT DO NOME DO PROJETO
    nome_projeto = st.text_input("Nome do Projeto / Obra", placeholder="Ex: Ed. Diogenes e Kely")
    
    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload is not None and nome_projeto:
        st.info(f"Arquivo: {arquivo_upload.name} | Obra: {nome_projeto}")
        
        if st.button("🚀 PROCESSAR E ADICIONAR", type="primary"):
            try:
                # Salva IFC temporário
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                # Processa Dados
                with st.spinner('Lendo IFC e extraindo dados...'):
                    novos_dados = processar_ifc(caminho_temp, nome_projeto)
                
                # Lógica Inteligente de Banco de Dados
                with st.spinner('Sincronizando com Google Sheets...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    ws = sh.sheet1
                    
                    # 1. Baixa tudo que já tem lá
                    dados_existentes = ws.get_all_records()
                    df_antigo = pd.DataFrame(dados_existentes)
                    
                    # 2. Se já tem dados, remove se houver duplicata deste mesmo projeto
                    # (Isso previne duplicidade se você subir o mesmo arquivo duas vezes)
                    if not df_antigo.empty and 'Projeto' in df_antigo.columns:
                        # Mantém tudo que NÃO é do projeto atual
                        df_limpo = df_antigo[df_antigo['Projeto'] != nome_projeto]
                    else:
                        df_limpo = pd.DataFrame()

                    # 3. Junta o Antigo Limpo + O Novo
                    df_novo = pd.DataFrame(novos_dados)
                    df_final = pd.concat([df_limpo, df_novo], ignore_index=True)
                    
                    # 4. Sobe tudo de volta
                    ws.clear()
                    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())
                
                # Gera PDF
                with st.spinner('Gerando arquivo de Etiquetas...'):
                    pdf_buffer = gerar_pdf_memoria(novos_dados, nome_projeto)
                
                st.success(f"✅ Projeto '{nome_projeto}' atualizado com sucesso!")
                st.metric(label="Total de Pilares na Base", value=len(df_final), delta=len(novos_dados))
                
                st.download_button("📥 BAIXAR ETIQUETAS (PDF)", pdf_buffer, f"Etiquetas_{nome_projeto}.pdf", "application/pdf")
                os.remove(caminho_temp)
                
            except Exception as e:
                st.error(f"Erro durante o processamento: {e}")
                
    elif arquivo_upload and not nome_projeto:
        st.warning("⚠️ Atenção: Você precisa digitar o nome do projeto antes de processar.")

if __name__ == "__main__":
    main()
