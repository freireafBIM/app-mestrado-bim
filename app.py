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

# --- VARIAVEL GLOBAL PARA CACHE DE ARMADURA ---
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
        return None

    client = gspread.authorize(creds)
    return client

def limpar_string(texto):
    if not texto: return "X"
    return "".join(e for e in str(texto) if e.isalnum()).upper()

# --- LÓGICA DE EXTRAÇÃO BLINDADA (NOVA) ---

def indexar_todas_armaduras(ifc_file):
    """
    Lê TODAS as barras do arquivo (sem depender de vínculo com pilar).
    Procura no nome da barra a quem ela pertence (Ex: '1 P4 ...').
    """
    global CACHE_ARMADURAS_POR_NOME
    CACHE_ARMADURAS_POR_NOME = {}
    
    # Pega todas as barras do projeto
    barras = ifc_file.by_type('IfcReinforcingBar')
    
    for bar in barras:
        nome_completo = bar.Name # Ex: "1 P1 \X\D810.00 C=280.00"
        if not nome_completo: continue
        
        # Regex para achar o dono: Procura "P" seguido de números (ex: P1, P12)
        # O padrão TQS geralmente é: Quantidade + Espaço + NOME + Espaço
        match_nome = re.search(r'^\d+\s+(P\d+)\s+', nome_completo)
        
        if match_nome:
            nome_pilar = match_nome.group(1) # "P1"
            
            # Extrai Bitola do texto (ex: \X\D810.00)
            match_bitola = re.search(r'\\X\\D8\s*([0-9\.]+)', nome_completo)
            bitola = 0.0
            
            if match_bitola:
                bitola = float(match_bitola.group(1))
            elif hasattr(bar, "NominalDiameter") and bar.NominalDiameter:
                bitola = bar.NominalDiameter * 1000 # Converte m para mm
            
            if bitola > 0:
                # Inicializa a lista se não existir
                if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
                    CACHE_ARMADURAS_POR_NOME[nome_pilar] = []
                
                # Adiciona a bitola à lista desse pilar
                CACHE_ARMADURAS_POR_NOME[nome_pilar].append(bitola)

def obter_armadura_do_cache(nome_pilar):
    """Busca a armadura no dicionário global criado anteriormente."""
    if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
        return "Sem armadura vinculada (Verificar TXT do TQS)"
    
    lista_bitolas = CACHE_ARMADURAS_POR_NOME[nome_pilar]
    c = Counter(lista_bitolas)
    
    # Ordena decrescente (mais grossa primeiro)
    return " + ".join([f"{qtd} ø{diam:.1f}" for diam, qtd in sorted(c.items(), reverse=True)])

def extrair_secao_geometria_pura(pilar):
    """
    Calcula o Bounding Box (Caixa Envolvente) varrendo todos os pontos 3D.
    Funciona para Extrusão, Malha (Mesh) e MappedItem.
    """
    # 1. Tenta Psets TQS primeiro (Mais rápido)
    psets = ifcopenshell.util.element.get_psets(pilar)
    if 'TQS_Geometria' in psets:
        d = psets['TQS_Geometria']
        b = d.get('Dimensao_b1') or d.get('B')
        h = d.get('Dimensao_h1') or d.get('H')
        if b and h:
            vals = sorted([float(b), float(h)])
            if vals[0] < 3.0: vals = [v*100 for v in vals] # Ajuste metros
            return f"{vals[0]:.0f}x{vals[1]:.0f}"

    # 2. Varredura 3D (Fallback para quando Pset falha)
    pontos_x = []
    pontos_y = []

    def coletar_pontos(item):
        # Se for um ponto cartesiano
        if item.is_a('IfcCartesianPoint'):
            coord = item.Coordinates
            if len(coord) >= 2:
                pontos_x.append(coord[0])
                pontos_y.append(coord[1])
        
        # Se for Malha (TQS usa muito isso)
        elif item.is_a('IfcFaceBasedSurfaceModel'):
            for face_set in item.FbsmFaces:
                coletar_pontos(face_set)
        elif item.is_a('IfcConnectedFaceSet'):
            for face in item.CfsFaces:
                coletar_pontos(face)
        elif item.is_a('IfcFace'):
            for bound in item.Bounds:
                coletar_pontos(bound)
        elif item.is_a('IfcFaceBound'):
            coletar_pontos(item.Bound)
        elif item.is_a('IfcPolyLoop'):
            for pt in item.Polygon:
                coletar_pontos(pt)
        
        # Se for Extrusão padrão
        elif item.is_a('IfcExtrudedAreaSolid'):
            coletar_pontos(item.SweptArea)
        elif item.is_a('IfcRectangleProfileDef'):
            # Para retângulo, calcula os cantos teóricos
            x = item.XDim / 2
            y = item.YDim / 2
            pontos_x.extend([x, -x])
            pontos_y.extend([y, -y])
        
        # Se for Mapeado (Cópia de outro pilar)
        elif item.is_a('IfcMappedItem'):
            coletar_pontos(item.MappingSource.MappedRepresentation)
        elif item.is_a('IfcShapeRepresentation'):
            for i in item.Items:
                coletar_pontos(i)

    # Executa a varredura
    if pilar.Representation:
        for rep in pilar.Representation.Representations:
            if rep.RepresentationIdentifier in ['Body', 'Mesh']:
                for item in rep.Items:
                    coletar_pontos(item)

    # Calcula distâncias
    if pontos_x and pontos_y:
        largura = max(pontos_x) - min(pontos_x)
        altura = max(pontos_y) - min(pontos_y)
        
        # Converte para cm se estiver em metros (TQS usa metros)
        if largura < 3.0: largura *= 100
        if altura < 3.0: altura *= 100
        
        dims = sorted([largura, altura])
        # Arredonda para inteiro mais próximo (evita 19.9999)
        return f"{dims[0]:.0f}x{dims[1]:.0f}"

    return "N/A"

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    
    # --- PASSO CRÍTICO: Ler todas as armaduras do arquivo antes ---
    indexar_todas_armaduras(ifc_file)
    # --------------------------------------------------------------
    
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
        
        # ID Único Robusto
        id_unico_pilar = f"{sufixo_nome}-{guid}-{sufixo_pav}-{id_projeto_input}"

        # Extração de Seção (Nova lógica de varredura)
        secao = extrair_secao_geometria_pura(pilar)
        
        # Extração de Armadura (Nova lógica de busca por nome)
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
        if st.button("🚀 PROCESSAR DADOS", type="primary"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_upload.getvalue())
                    caminho_temp = tmp_file.name
                
                with st.spinner('Lendo IFC (Geometria e Armaduras)...'):
                    dados_pilares = processar_ifc(caminho_temp, id_projeto)
                os.remove(caminho_temp)

                with st.spinner('Atualizando Google Sheets...'):
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
                
                st.success(f"✅ Projeto processado! {len(dados_pilares)} pilares identificados.")
                st.download_button("📥 BAIXAR PDF", pdf_buffer, f"Etiquetas_{id_projeto}.pdf", "application/pdf")
                
            except Exception as e:
                st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
