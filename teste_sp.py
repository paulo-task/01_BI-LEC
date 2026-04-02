import requests
from urllib.parse import urlparse
from dotenv import load_dotenv
import os

# Carrega as variáveis do arquivo .pass
try:
    load_dotenv(dotenv_path=".pass")
    print("✅ Credenciais carregadas do .pass")
except:
    print("⚠️ Arquivo .pass não encontrado")

# Pega as variáveis
SP_CLIENT_ID = os.getenv("SP_CLIENT_ID")
SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
SP_TENANT_ID = os.getenv("SP_TENANT_ID")
SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"

if not SP_CLIENT_ID:
    print("❌ SP_CLIENT_ID não encontrado!")
    exit(1)

print("1. Obtendo token...")
url_token = f'https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token'
data = {
    'grant_type': 'client_credentials',
    'client_id': SP_CLIENT_ID,
    'client_secret': SP_CLIENT_SECRET,
    'scope': 'https://graph.microsoft.com/.default'
}
r = requests.post(url_token, data=data)

if r.status_code != 200:
    print(f"❌ Erro ao obter token: {r.status_code}")
    exit(1)

token = r.json()['access_token']
print("✅ Token obtido")

print("2. Obtendo Site ID...")
parsed = urlparse(SITE_URL)
host = parsed.netloc
site_path = parsed.path.strip("/")
url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
headers = {"Authorization": f"Bearer {token}"}
r = requests.get(url_site, headers=headers)

if r.status_code != 200:
    print(f"❌ Erro ao obter Site ID: {r.status_code}")
    exit(1)

site_id = r.json()["id"]
print(f"✅ Site ID: {site_id}")

print("3. Buscando a biblioteca Workspace...")
url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
r = requests.get(url_drives, headers={"Authorization": f"Bearer {token}"})

drive_id = None
if r.status_code == 200:
    drives = r.json().get('value', [])
    for drive in drives:
        print(f"   - {drive.get('name')}: {drive.get('webUrl')}")
        if drive.get('name') == 'Workspace':
            drive_id = drive.get('id')
            print(f"\n✅ Usando biblioteca: Workspace (ID: {drive_id})")
            break

if not drive_id:
    print("❌ Biblioteca 'Workspace' não encontrada!")
    exit(1)

print("4. Enviando arquivo TXT para a biblioteca Workspace...")
pasta = "BI_LEC/01_ELF_Diario"
nome_arquivo = "teste_workspace_correto.txt"
conteudo = "Teste usando biblioteca Workspace - " + str(__import__('datetime').datetime.now())

# Usa a biblioteca Workspace
url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta}/{nome_arquivo}:/content"
print(f"URL de upload: {url_upload}")

headers_upload = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "text/plain"
}
r = requests.put(url_upload, headers=headers_upload, data=conteudo.encode('utf-8'))

if r.status_code in [200, 201]:
    print("✅ Upload realizado com sucesso!")
    print(f"URL do arquivo: {r.json().get('webUrl')}")
    
    # Verifica se o caminho está correto
    web_url = r.json().get('webUrl')
    if "Workspace/BI_LEC/01_ELF_Diario" in web_url:
        print("✅ Caminho está CORRETO!")
    else:
        print(f"⚠️ Caminho: {web_url}")
else:
    print(f"❌ Erro: {r.status_code}")
    print(r.text)