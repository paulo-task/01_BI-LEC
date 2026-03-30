print("--- Script iniciando via Graph API (Biblioteca Independente) ---")

import os
import requests

CLIENT_ID = os.environ.get("SP_CLIENT_ID")
CLIENT_SECRET = os.environ.get("SP_CLIENT_SECRET")
TENANT_ID = os.environ.get("SP_TENANT_ID") 

SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG" 
NOME_BIBLIOTECA = "Workspace"  # O nome exato da sua biblioteca independente

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json().get("access_token")

def get_site_id(token):
    from urllib.parse import urlparse
    parsed = urlparse(SITE_URL)
    host = parsed.netloc
    site_path = parsed.path.lstrip("/") 
    url = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()["id"]

def get_workspace_drive_id(token, site_id):
    """Busca o ID da biblioteca específica chamada 'Workspace'"""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    
    drives = r.json().get("value", [])
    for drive in drives:
        if drive['name'] == NOME_BIBLIOTECA:
            print(f"✅ Biblioteca '{NOME_BIBLIOTECA}' encontrada!")
            return drive['id']
    
    raise Exception(f"❌ Biblioteca '{NOME_BIBLIOTECA}' não encontrada no site.")

def upload_arquivo(token, drive_id):
    print("--- Iniciando o upload do arquivo ---")
    # Agora o caminho começa direto de dentro da Workspace
    PASTA_ALVO = "BI_LEC/01_ELF_Diario" 
    NOME_ARQUIVO = "PROVA_FINAL_BIBLIOTECA.txt"
    CONTEUDO = b"Teste de gravacao em biblioteca independente ok"

    # Usamos o drive_id da Workspace em vez de 'root'
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{PASTA_ALVO}/{NOME_ARQUIVO}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "text/plain",
    }
    r = requests.put(url, headers=headers, data=CONTEUDO)
    r.raise_for_status()
    print(f"💎 SUCESSO! Arquivo enviado para a Workspace. Link: {r.json().get('webUrl')}")

if __name__ == "__main__":
    try:
        token = get_access_token()
        site_id = get_site_id(token)
        drive_id = get_workspace_drive_id(token, site_id)
        upload_arquivo(token, drive_id)
    except Exception as e:
        print(f"❌ ERRO: {e}")
