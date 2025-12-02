import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import streamlit as st # Só para ler os secrets se precisar

# CONFIGURAÇÕES
ARQUIVO_CREDENCIAIS = "credenciais.json"

# COLE AQUI O ID QUE VOCÊ ESTÁ TENTANDO USAR
ID_PASTA_TESTE = "1I37hXwx6zpIGItxpM_guTQFEls-W8gff" 

def main():
    print("--- INICIANDO DIAGNÓSTICO ---")
    
    # 1. Tenta carregar credenciais
    scopes = ['https://www.googleapis.com/auth/drive']
    
    try:
        # Tenta pegar dos Secrets (Streamlit Cloud) ou Local
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            print("Usando credenciais do Streamlit Secrets.")
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        elif os.path.exists(ARQUIVO_CREDENCIAIS):
            print("Usando arquivo credenciais.json local.")
            creds = Credentials.from_service_account_file(ARQUIVO_CREDENCIAIS, scopes=scopes)
        else:
            print("ERRO: Nenhuma credencial encontrada!")
            return
            
        service = build('drive', 'v3', credentials=creds)
        
        # 2. Descobre QUEM é o Robô
        about = service.about().get(fields="user").execute()
        email_robo = about['user']['emailAddress']
        print(f"\n[1] O Python está usando este e-mail:")
        print(f"    >>> {email_robo} <<<")
        print("    (Vá no Google Drive, clique na pasta > Compartilhar e veja se ESTE e-mail exato está lá como EDITOR)")
        
        # 3. Tenta ENXERGAR a pasta
        print(f"\n[2] Tentando acessar a pasta ID: {ID_PASTA_TESTE}")
        try:
            folder = service.files().get(fileId=ID_PASTA_TESTE, fields="name, capabilities").execute()
            print(f"    SUCESSO! O Robô encontrou a pasta: '{folder.get('name')}'")
            
            # Verifica se pode escrever
            pode_editar = folder['capabilities']['canAddChildren']
            print(f"    Pode gravar arquivos nela? {'SIM' if pode_editar else 'NÃO (Verifique se é Editor)'}")
            
        except Exception as e:
            print(f"    FALHA: O Robô NÃO consegue ver a pasta.")
            print(f"    Motivo provável: Você compartilhou a pasta com outro e-mail, não com o {email_robo}")
            print(f"    Erro técnico: {e}")

    except Exception as e:
        print(f"Erro fatal no script: {e}")

if __name__ == "__main__":
    main()
