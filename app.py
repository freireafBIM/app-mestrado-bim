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
import re # Biblioteca de Expressões Regulares (para ler o texto do TQS)

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- FUNÇÕES DE CONEXÃO ---
def conectar_google_sheets():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Tenta conectar via Secrets (Nuvem) ou Arquivo Local
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif os.path.exists(ARQUIVO_CREDENCIAIS):
        creds = Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
    else:
        st.error("ERRO CRÍTICO: Credenciais não encontradas (JSON ou Secrets).")
        st.stop()
        return None

    client = gspread.authorize(creds)
    return client

# --- FUNÇÕES AUXILIARES ---
def limpar_string(texto):
    """Remove caracteres especiais para criar IDs seguros."""
    if not texto: return "X"
    # Mantém apenas letras e números e converte para maiúsculo
    return "".join(e for e in str(texto) if e.isalnum()).upper()

# --- LÓGICA DE EXTRAÇÃO  (TQS) ---

def extrair_texto_armadura(pilar):
    """
    Lê as armaduras exportadas pelo TQS como objetos IfcReinforcingBar.
    Interpreta o texto '1 N10 ø10.0 C=...' gravado no nome da barra.
    """
    # Dicionário para somar as barras: { '10.0': 4, '5.0': 12 }
    contagem_bitolas = {}
    
    # 1. Busca objetos de armadura conectados ao pilar
    # O TQS conecta as barras ao pilar através da relação 'IfcRelAggregates'
    relacoes = getattr(pilar, 'IsDecomposedBy', [])
    
    encontrou_barras = False
    
    for rel in relacoes:
        if rel.is_a('IfcRelAggregates'):
            for obj in rel.RelatedObjects:
                if obj.is_a('IfcReinforcingBar'):
                    encontrou_barras = True
                    nome_bruto = obj.Name # Ex: "1 P4 \X\D85.00 C=84.00"
                    
                    # Tenta extrair o diâmetro usando "Regex" (busca padrões de texto)
                    # Procura por "\X\D8" seguido de números (ex: 5.00 ou 10.00)
                    match = re.search(r'\\X\\D8\s*([0-9\.]+)', nome_bruto)
                    
                    if match:
                        diametro = float(match.group(1)) # Pega o número (5.00)
                        # Arredonda para ficar bonito (10.0, 12.5, 5.0)
                        diametro_str = f"{diametro:.1f}" 
                        
                        # Adiciona na contagem
                        qt_atual = contagem_bitolas.get(diametro_str, 0)
                        contagem_bitolas[diametro_str] = qt_atual + 1
                    else:
                        # Se não achou no nome, tenta pegar da propriedade física (fallback)
                        if hasattr(obj, "NominalDiameter"):
                             # IFC usa metros, converte para mm
                            d = round(obj.NominalDiameter * 1000, 1)
                            d_str = f"{d:.1f}"
                            contagem_bitolas[d_str] = contagem_bitolas.get(d_str, 0) + 1

    # 2. Formata o texto final (Ex: "4 ø10.0 + 12 ø5.0")
    if not encontrou_barras:
        # Tenta ler Psets de texto se não tiver objetos 3D (Plano C)
        psets = ifcopenshell.util.element.get_psets(pilar)
        for nome, dados in psets.items():
            if 'Armadura' in nome or 'Reinforcement' in nome:
                for k, v in dados.items():
                    if isinstance(v, str) and len(v) > 5: return v
        return "Verificar Detalhamento (Sem barras vinculadas)"
    
    # Monta a string resumo ordenando por bitola (grossas primeiro)
    textos = []
    # Ordena as bitolas (converter para float para ordenar corretamente: 10.0 > 5.0)
    bitolas_ordenadas = sorted(contagem_bitolas.keys(), key=lambda x: float(x), reverse=True)
    
    for diam in bitolas_ordenadas:
        qtd = contagem_bitolas[diam]
        textos.append(f"{qtd} ø{diam}")
        
    return " + ".join(textos)

def extrair_secao_robusta(pilar):
    """
    Tenta extrair a seção (Ex: 14x30) priorizando dados do TQS
    e usando geometria bruta como backup.
    """
    psets = ifcopenshell.util.element.get_psets(pilar)
    b_val, h_val = 0, 0
    encontrou = False

    # 1. TENTATIVA TQS (Pset TQS_Geometria)
    if 'TQS_Geometria' in psets:
        dados = psets['TQS_Geometria']
        # TQS usa 'Dimensao_b1' e 'Dimensao_h1'
        b = dados.get('Dimensao_b1') or dados.get('B')
        h = dados.get('Dimensao_h1') or dados.get('H')
        if b and h:
            b_val, h_val = float(b), float(h)
            encontrou = True

    # 2. TENTATIVA GENÉRICA (ExtrudedAreaSolid)
    if not encontrou and pilar.Representation:
        for rep in pilar.Representation.Representations:
            if rep.RepresentationIdentifier == 'Body':
                for item in rep.Items:
                    if item.is_a('IfcExtrudedAreaSolid'):
                        perfil = item.SweptArea
                        if perfil.is_a('IfcRectangleProfileDef'):
                            b_val = perfil.XDim
                            h_val = perfil.YDim
                            encontrou = True

    if encontrou:
        # Tratamento de Unidades: Se for pequeno (< 4.0), assume metros e converte p/ cm
        if b_val < 4.0: b_val *= 100
        if h_val < 4.0: h_val *= 100
        
        # Ordena sempre Menor x Maior para padronizar (Ex: 14x30)
        dims = sorted([b_val, h_val])
        return f"{dims[0]:.0f}x{dims[1]:.0f}"
    
    return "N/A"

def processar_ifc(caminho_arquivo, id_projeto_input):
    """Lê o IFC e gera a lista de dados com Chaves Primárias Robustas."""
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        # Identifica Pavimento
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name
        
        # Limpa textos para usar no ID
        sufixo_pav = limpar_string(pavimento)
        sufixo_nome = limpar_string(nome)

        # --- CHAVE PRIMÁRIA ROBUSTA (A CORREÇÃO PRINCIPAL) ---
        # Formato: NOME - GUID - PAVIMENTO - PROJETO
        # Isso impede que P1 seja confundido com P12 ou P1 de outro andar
        id_unico_pilar = f"{sufixo_nome}-{guid}-{sufixo_pav}-{id_projeto_input}"

        secao = extrair_secao_robusta(pilar)
        armadura = extrair_texto_armadura(pilar)

        dados.append({
            'ID_Unico': id_unico_pilar,      # Chave Única (Key)
            'Projeto_Ref': id_projeto_input, # Chave Estrangeira (Ref)
            'Nome': nome,                    # Label
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    # Ordena visualmente por Pavimento e Nome
    dados.sort(key=lambda x: (x['Pavimento'], x['Nome']))
    return dados

def gerar_pdf_memoria(dados_pilares, nome_projeto_legivel):
    """Gera o PDF com etiquetas na memória RAM."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largura_pag, altura_pag = A4
    largura_etq, altura_etq = 90*mm, 50*mm
    margem, espaco = 10*mm, 5*mm
    
    x, y = margem, altura_pag - margem - altura_etq
    
    for pilar in dados_pilares:
        c.setLineWidth(0.5)
        c.rect(x, y, largura_etq, altura_etq)
        
        # Gera QR Code contendo o ID ÚNICO ROBUSTO
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

# --- FRONTEND (INTERFACE) ---

def main():
    st.set_page_config(page_title="Sistema Conferência Armaduras", page_icon="🏗️")
    
    # Login
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

    st.title("🏗️ Sistema Conferência Armaduras")
    
    # Inputs
    nome_projeto_legivel = st.text_input("Nome da Obra (Legível)", placeholder="Ex: Edifício Diogenes")
    # Cria ID técnico (Ex: EDIFICIODIOGENES)
    id_projeto = limpar_string(nome_projeto_legivel)
    
    if nome_projeto_legivel:
        st.caption(f"ID Técnico do Projeto: {id_projeto}")

    arquivo_upload = st.file_uploader("Carregar arquivo IFC", type=["ifc"])
    
    if arquivo_upload and nome_projeto_legivel:
        if st.button("🚀 PROCESSAR E SINCRONIZAR", type="primary"):
            try:
                # Salva IFC temporário
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                # 1. Processar Pilares (Extração)
                with st.spinner('Lendo IFC e gerando IDs Únicos...'):
                    dados_pilares = processar_ifc(caminho_temp, id_projeto)
                os.remove(caminho_temp)

                # 2. Atualizar Google Sheets
                with st.spinner('Sincronizando Tabelas (Projetos e Pilares)...'):
                    client = conectar_google_sheets()
                    sh = client.open(NOME_PLANILHA_GOOGLE)
                    
                    # --- TABELA PROJETOS ---
                    try: ws_proj = sh.worksheet("Projetos")
                    except: ws_proj = sh.add_worksheet("Projetos", 100, 5)
                    
                    lista_proj = ws_proj.get_all_records()
                    df_proj = pd.DataFrame(lista_proj)
                    # Remove duplicata do projeto atual
                    if not df_proj.empty and 'ID_Projeto' in df_proj.columns:
                        df_proj = df_proj[df_proj['ID_Projeto'] != id_projeto]
                    
                    novo_proj = {
                        'ID_Projeto': id_projeto,
                        'Nome_Obra': nome_projeto_legivel,
                        'Data_Upload': datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                        'Total_Pilares': len(dados_pilares)
                    }
                    df_proj_final = pd.concat([df_proj, pd.DataFrame([novo_proj])], ignore_index=True)
                    ws_proj.clear()
                    ws_proj.update([df_proj_final.columns.values.tolist()] + df_proj_final.values.tolist())

                    # --- TABELA PILARES ---
                    try: ws_pil = sh.worksheet("Pilares")
                    except: ws_pil = sh.add_worksheet("Pilares", 1000, 10)
                    
                    lista_pil = ws_pil.get_all_records()
                    df_pil = pd.DataFrame(lista_pil)
                    
                    # Remove APENAS os pilares deste projeto específico (Limpeza)
                    if not df_pil.empty and 'Projeto_Ref' in df_pil.columns:
                        df_pil = df_pil[df_pil['Projeto_Ref'] != id_projeto]
                    
                    df_pil_novos = pd.DataFrame(dados_pilares)
                    df_pil_final = pd.concat([df_pil, df_pil_novos], ignore_index=True)
                    
                    ws_pil.clear()
                    ws_pil.update([df_pil_final.columns.values.tolist()] + df_pil_final.values.tolist())

                # 3. PDF
                with st.spinner('Gerando Etiquetas PDF...'):
                    pdf_buffer = gerar_pdf_memoria(dados_pilares, nome_projeto_legivel)
                
                st.success(f"✅ Projeto '{nome_projeto_legivel}' atualizado com sucesso!")
                st.download_button("📥 BAIXAR PDF DAS ETIQUETAS", pdf_buffer, f"Etiquetas_{id_projeto}.pdf", "application/pdf")
                
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()


