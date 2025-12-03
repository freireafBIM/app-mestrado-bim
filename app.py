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
from reportlab.lib import colors
import io
import datetime
import re
from collections import Counter

# --- CONFIGURAÇÕES ---
ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_PLANILHA_GOOGLE = "Sistema_Conferencia_BIM"

# --- CACHE GLOBAL PARA ARMADURAS ---
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

# --- LÓGICA DE EXTRAÇÃO DE ARMADURA (TQS) ---

def indexar_todas_armaduras(ifc_file):
    """
    Lê todas as barras do arquivo IFC e cria um dicionário agrupando por Pilar.
    Padrão TQS detectado: '1 P1 \X\D810.00 C=280.00'
    """
    global CACHE_ARMADURAS_POR_NOME
    CACHE_ARMADURAS_POR_NOME = {}
    
    # Busca todas as barras, independente de onde estejam na hierarquia
    barras = ifc_file.by_type('IfcReinforcingBar')
    
    for bar in barras:
        nome_completo = bar.Name 
        if not nome_completo: continue
        
        # 1. Identificar o DONO da barra (ex: P1, P12)
        # Regex: Procura um número, espaço, Letra P seguida de números, espaço
        # Ex: "1 P1 " -> captura P1
        match_nome = re.search(r'^\d+\s+(P\d+)\s', nome_completo)
        
        if match_nome:
            nome_pilar = match_nome.group(1) # Ex: "P1"
            bitola = 0.0
            
            # 2. Identificar a BITOLA (Diâmetro)
            # Tenta achar o padrão do TQS (\X\D8) ou símbolo unicode (Ø)
            # Pega o primeiro número decimal que aparece DEPOIS do nome do pilar
            resto_string = nome_completo.split(nome_pilar, 1)[1]
            match_bitola = re.search(r'([0-9]+\.[0-9]+)', resto_string)
            
            if match_bitola:
                bitola = float(match_bitola.group(1))
            
            # Fallback: Se não achou no texto, pega da propriedade física do objeto
            elif hasattr(bar, "NominalDiameter") and bar.NominalDiameter:
                bitola = bar.NominalDiameter * 1000 # Converte m para mm
            
            if bitola > 0:
                if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
                    CACHE_ARMADURAS_POR_NOME[nome_pilar] = []
                CACHE_ARMADURAS_POR_NOME[nome_pilar].append(bitola)

def obter_armadura_do_cache(nome_pilar):
    """Retorna o resumo textual da armadura para um pilar específico."""
    if nome_pilar not in CACHE_ARMADURAS_POR_NOME:
        return "Verificar Detalhamento"
    
    lista_bitolas = CACHE_ARMADURAS_POR_NOME[nome_pilar]
    c = Counter(lista_bitolas)
    
    # Ordena por bitola (do mais grosso para o mais fino)
    return " + ".join([f"{qtd} ø{diam:.1f}" for diam, qtd in sorted(c.items(), key=lambda item: item[0], reverse=True)])

# --- LÓGICA DE EXTRAÇÃO DE GEOMETRIA (UNIVERSAL) ---

def extrair_secao_universal(pilar):
    """
    Calcula o Bounding Box (Caixa Envolvente) varrendo todos os pontos 3D.
    Funciona para Extrusão, Malha (Mesh) e MappedItem do TQS.
    """
    # 1. Tenta Psets TQS primeiro (Mais rápido/preciso se existir)
    psets = ifcopenshell.util.element.get_psets(pilar)
    if 'TQS_Geometria' in psets:
        d = psets['TQS_Geometria']
        b = d.get('Dimensao_b1') or d.get('B')
        h = d.get('Dimensao_h1') or d.get('H')
        if b and h:
            vals = sorted([float(b), float(h)])
            if vals[0] < 3.0: vals = [v*100 for v in vals] # Ajuste metros -> cm
            return f"{vals[0]:.0f}x{vals[1]:.0f}"

    # 2. Varredura 3D (Fallback para geometria complexa)
    if not pilar.Representation: return "N/A"
    
    pontos_x = []
    pontos_y = []
    
    def coletar_pontos(item):
        # Proteção contra tipos inválidos
        if item is None: return
        if isinstance(item, (list, tuple)):
            for i in item: coletar_pontos(i)
            return
        if not hasattr(item, 'is_a'): return

        # Se achou um ponto cartesiano
        if item.is_a('IfcCartesianPoint'):
            if hasattr(item, 'Coordinates') and len(item.Coordinates) >= 2:
                pontos_x.append(item.Coordinates[0])
                pontos_y.append(item.Coordinates[1])
            return

        # Recursividade para atributos comuns de geometria
        atributos = [
            'Points', 'OuterCurve', 'PolygonalBoundary', 
            'FbsmFaces', 'CfsFaces', 'Bounds', 'Bound', 
            'Items', 'MappingSource', 'MappedRepresentation', 
            'Polygon', 'SweptArea'
        ]
        for attr in atributos:
            if hasattr(item, attr):
                coletar_pontos(getattr(item, attr))

    # Inicia a varredura
    for rep in pilar.Representation.Representations:
        if rep.RepresentationIdentifier in ['Body', 'Mesh', 'Box', 'Facetation']:
            for item in rep.Items:
                coletar_pontos(item)
    
    # Calcula Bounding Box
    if pontos_x and pontos_y:
        try:
            largura = max(pontos_x) - min(pontos_x)
            altura = max(pontos_y) - min(pontos_y)
            
            if largura < 0.01 or altura < 0.01: return "N/A" # Ignora pontos zerados

            # Converte para cm se estiver em metros
            if largura < 3.0: largura *= 100
            if altura < 3.0: altura *= 100
            
            dims = sorted([largura, altura])
            return f"{dims[0]:.0f}x{dims[1]:.0f}"
        except:
            return "N/A"

    return "N/A"

def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    
    # PASSO 1: Indexar todas as barras antes de ler os pilares
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
        
        # ID ÚNICO COMPOSTO (Evita duplicidade)
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

# --- GERAÇÃO DE PDF (REVISADO E FORMATADO) ---

def gerar_pdf_memoria(dados_pilares, nome_projeto_legivel):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largura_pag, altura_pag = A4
    
    # Configurações da Etiqueta
    largura_etq = 90 * mm
    altura_etq = 50 * mm
    margem_x = 10 * mm
    margem_y = 10 * mm # Margem do topo
    espaco_x = 5 * mm
    espaco_y = 5 * mm
    
    # Posição inicial (Canto superior esquerdo)
    x = margem_x
    y = altura_pag - margem_y - altura_etq
    
    c.setTitle(f"Etiquetas - {nome_projeto_legivel}")
    
    for i, pilar in enumerate(dados_pilares):
        # 1. Desenha Borda da Etiqueta
        c.setLineWidth(1)
        c.setStrokeColor(colors.black)
        c.rect(x, y, largura_etq, altura_etq)
        
        # 2. Gera QR Code
        qr = qrcode.QRCode(box_size=10, border=0) # Border 0 para caber melhor
        qr.add_data(pilar['ID_Unico'])
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        temp_qr_path = f"temp_{i}.png"
        img_qr.save(temp_qr_path)
        
        # Desenha QR Code (Quadrado de 35mm)
        c.drawImage(temp_qr_path, x + 3*mm, y + 7.5*mm, width=35*mm, height=35*mm)
        os.remove(temp_qr_path)
        
        # 3. Textos (Lado Direito)
        texto_x = x + 42*mm
        
        # Título (Nome do Pilar)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(texto_x, y + 38*mm, f"PILAR: {pilar['Nome']}")
        
        # Seção
        c.setFont("Helvetica", 12)
        c.drawString(texto_x, y + 30*mm, f"Seção: {pilar['Secao']}")
        
        # Pavimento (Importante estar grande)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(texto_x, y + 22*mm, f"{pilar['Pavimento']}")
        
        # Nome da Obra (Menor, no rodapé)
        c.setFont("Helvetica-Oblique", 8)
        c.setFillColor(colors.gray)
        c.drawString(texto_x, y + 8*mm, f"Obra: {nome_projeto_legivel[:18]}...")
        c.setFillColor(colors.black)
        
        # 4. Linha de Corte (Pontilhada ao redor)
        c.setDash(3, 3)
        c.setLineWidth(0.2)
        c.rect(x-
