import os
import time
import io
import re
import sys
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis
try:
    load_dotenv(dotenv_path=".pass")
except:
    pass

usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

if is_github_actions():
    USE_HEADLESS = True
else:
    USE_HEADLESS = False

# Mapeamento dos relatórios
MAPEAMENTO = {
    "Efetividade de Leitura Faturamento": {
        "pasta_sp": "BI_LEC/01_ELF_Diario",
        "nome_base": "EFL",
        "regra": "data_dia_unico"
    },
    "Produtividade Diária Leiturista - Analítico": {
        "pasta_sp": "BI_LEC/02_PDL_Analitico",
        "nome_base": "Produtividade Diária Leiturista - Analítico",
        "regra": "data_dia"
    },
    "Instalações Não Visitadas (Diário)": {
        "pasta_sp": "BI_LEC/04_N_Visitado_Diario",
        "nome_base": "Instalações Não Visitadas",
        "regra": "data_dia"
    },
    "Instalações Não Visitadas (Histórico)": {
        "pasta_sp": "BI_LEC/05_N_Visitado_Historico",
        "nome_base": "Instalações Não Visitadas",
        "regra": "data_dia"
    }
}

def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo para o SharePoint"""
    if not is_github_actions():
        return True
    
    try:
        import requests
        from urllib.parse import urlparse
        
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID")
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
        SP_TENANT_ID = os.getenv("SP_TENANT_ID")
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        # Obter token
        url_token = f'https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token'
        data = {
            'grant_type': 'client_credentials',
            'client_id': SP_CLIENT_ID,
            'client_secret': SP_CLIENT_SECRET,
            'scope': 'https://graph.microsoft.com/.default'
        }
        r = requests.post(url_token, data=data)
        r.raise_for_status()
        token = r.json()['access_token']
        
        # Obter Site ID
        parsed = urlparse(SITE_URL)
        host = parsed.netloc
        site_path = parsed.path.strip("/")
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url_site, headers=headers)
        r.raise_for_status()
        site_id = r.json()["id"]
        
        # Obter Drive Workspace
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        r = requests.get(url_drives, headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        
        drive_id = None
        for drive in r.json().get('value', []):
            if drive.get('name') == 'Workspace':
                drive_id = drive.get('id')
                break
        
        if not drive_id:
            raise Exception("Workspace não encontrado")
        
        # Upload
        url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content"
        headers_upload = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        r = requests.put(url_upload, headers=headers_upload, data=conteudo_bytes)
        r.raise_for_status()
        
        print(f"   ✅ Upload: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"   ❌ Erro upload: {e}")
        return False

def processar_zip(conteudo_zip, nome_base, regra, pasta_sp):
    """Processa o ZIP e envia para SharePoint"""
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo_zip)) as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()
            
            if regra == "data_dia_unico":
                nome_final = f"{nome_base} - {agora.strftime('%Y_%m_%d')}{ext}"
            elif regra == "data_dia":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m_%d')}{ext}"
            else:
                nome_final = f"{nome_base}{ext}"
            
            with z.open(arquivo_interno) as f:
                conteudo = f.read()
            
            return upload_to_sharepoint(conteudo, nome_final, pasta_sp)
            
    except Exception as e:
        print(f"   ❌ Erro processar ZIP: {e}")
        return False

def run(playwright: Playwright, busca=None) -> None:
    print(f"🏗️ Ambiente: {'GitHub' if is_github_actions() else 'Local'}")
    
    context = playwright.chromium.launch_persistent_context(
        "temp_navegador" if is_github_actions() else r"C:\Temp\chrome_profile",
        headless=USE_HEADLESS,
        args=['--no-sandbox'] if is_github_actions() else [],
        slow_mo=500,
    )
    page = context.pages[0]
    page.set_default_timeout(120000)

    try:
        print("🌐 Login...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        page.get_by_role("textbox", name="Usuário").fill(usuario)
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()
        page.wait_for_load_state("networkidle")
        
        print("📊 Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        
        # Aguarda tabela
        page.wait_for_selector("tbody tr", timeout=60000)
        time.sleep(3)
        
        linhas = page.locator("tbody tr").all()
        print(f"📋 {len(linhas)} linhas encontradas")
        
        # Baixar e enviar cada relatório
        import io
        
        for nome_rel, conf in MAPEAMENTO.items():
            if busca and not any(b.lower() in nome_rel.lower() for b in busca):
                continue
            
            print(f"\n🔍 {nome_rel}")
            
            # Encontrar linha mais recente
            maior_dt = None
            linha_alvo = None
            for linha in linhas:
                texto = linha.inner_text()
                if "Concluído" in texto:
                    m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
                    if m:
                        dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                        if maior_dt is None or dt > maior_dt:
                            maior_dt = dt
                            linha_alvo = linha
            
            if linha_alvo:
                print(f"   📥 Baixando...")
                with page.expect_download() as download_info:
                    linha_alvo.locator("button, a, img").first.click()
                
                download = download_info.value
                conteudo_zip = download.read()
                
                print(f"   📤 Enviando para SharePoint...")
                processar_zip(conteudo_zip, conf["nome_base"], conf["regra"], conf["pasta_sp"])
            else:
                print(f"   ⚠️ Nenhum relatório concluído")
        
        print("\n✅ Coletor finalizado!")
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        raise
    finally:
        context.close()

if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    with sync_playwright() as playwright:
        run(playwright, args)