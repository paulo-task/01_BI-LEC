import os
import time
import re
import sys
import io
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# =================================================================
# 1. CARREGAR CREDENCIAIS
# =================================================================
# Força a busca do arquivo .pass na mesma pasta deste script
diretorio_script = os.path.dirname(os.path.abspath(__file__))
caminho_pass = os.path.join(diretorio_script, ".pass")

if os.path.exists(caminho_pass):
    load_dotenv(dotenv_path=caminho_pass)
    print(f"✅ Arquivo .pass carregado de: {caminho_pass}")
else:
    load_dotenv()
    print("⚠️ Arquivo .pass não encontrado. Usando variáveis de ambiente.")

usuario = os.getenv("CPFL_USER", "")
senha = os.getenv("CPFL_PASS", "")

# Validação das credenciais
if not usuario or not senha:
    print("❌ ERRO: Credenciais CPFL não encontradas!")
    print("   Verifique o arquivo .pass ou as variáveis de ambiente")
    sys.exit(1)

print(f"✅ Credenciais CPFL carregadas: Usuário={usuario[:5]}...")

# =================================================================
# 2. DETECTAR AMBIENTE
# =================================================================
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

if is_github_actions():
    USE_HEADLESS = True
    print("🏗️ Ambiente: GitHub Actions (modo headless)")
else:
    USE_HEADLESS = False
    print("🖥️ Ambiente: Windows Local (modo visível)")

# =================================================================
# 3. MAPEAMENTO DOS RELATÓRIOS
# =================================================================
# Para SharePoint (GitHub)
MAPEAMENTO_SP = {
    "Efetividade de Leitura Faturamento": {
        "pasta_sp": "BI_LEC/01_ELF_Diario",
        "nome_base": "EFL",
        "regra": "data_dia_unico"
    },
    "Produtividade Diária Leiturista - Analítico": {
        "pasta_sp": "BI_LEC/02_PDL_Analitico",
        "nome_base": "Produtividade Diária Leiturista - Analítico",
        "regra": "data_dia"
    }
}

# Para pastas locais (PC)
PASTAS_LOCAIS = {
    "01_ELF_Diario": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
    "02_PDL_Analitico": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
    "04_N_Visitado_Diario": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario",
    "05_N_Visitado_Historico": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico"
}

# =================================================================
# 4. FUNÇÃO DE UPLOAD PARA SHAREPOINT (APENAS GITHUB)
# =================================================================
def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo para o SharePoint (apenas no GitHub Actions)"""
    try:
        import requests
        from urllib.parse import urlparse
        
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID", "").strip()
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "").strip()
        SP_TENANT_ID = os.getenv("SP_TENANT_ID", "").strip()
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        if not SP_CLIENT_ID:
            print("   ⚠️ SharePoint: credenciais não configuradas")
            return False
        
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
        
        print(f"   ☁️ Upload SharePoint: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"   ⚠️ Erro upload SharePoint: {e}")
        return False

# =================================================================
# 5. FUNÇÃO PARA SALVAR LOCALMENTE (PC)
# =================================================================
def salvar_localmente(conteudo_bytes, nome_arquivo, pasta_destino):
    """Salva arquivo localmente no PC"""
    try:
        os.makedirs(pasta_destino, exist_ok=True)
        caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
        
        with open(caminho_arquivo, 'wb') as f:
            f.write(conteudo_bytes)
        
        print(f"   💾 Salvo localmente: {caminho_arquivo}")
        return True
    except Exception as e:
        print(f"   ❌ Erro ao salvar localmente: {e}")
        return False

# =================================================================
# 6. PROCESSAR ZIP E SALVAR/ENVIAR
# =================================================================
def processar_zip(conteudo_zip, nome_base, regra, destino, tipo_destino="ambos"):
    """
    Processa o ZIP e salva/envia conforme o ambiente
    tipo_destino: "sharepoint", "local", "ambos"
    """
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
            
            sucesso = False
            
            # No GitHub: sempre tenta enviar para SharePoint
            if is_github_actions():
                sucesso = upload_to_sharepoint(conteudo, nome_final, destino)
            else:
                # No PC: salva localmente
                sucesso = salvar_localmente(conteudo, nome_final, destino)
            
            return sucesso
            
    except Exception as e:
        print(f"   ❌ Erro processar ZIP: {e}")
        return False

# =================================================================
# 7. FUNÇÃO PRINCIPAL
# =================================================================
def run(playwright: Playwright, busca=None) -> None:
    print(f"\n{'='*60}")
    print(f"🚀 INICIANDO COLETOR")
    print(f"{'='*60}")
    
    # Diretório para perfil do navegador
    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
        os.makedirs(user_dir, exist_ok=True)
    else:
        user_dir = r"C:\Temp\chrome_profile"
    
    # Configuração do navegador
    launch_args = ['--no-sandbox'] if is_github_actions() else []
    
    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,
        no_viewport=True,
        slow_mo=500,
    )
    page = context.pages[0]
    page.set_default_timeout(120000)

    try:
        # Login
        print("\n🌐 Fazendo login no portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        page.wait_for_load_state("networkidle")
        
        page.get_by_role("textbox", name="Usuário").fill(usuario)
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()
        page.wait_for_load_state("networkidle")
        print("✅ Login realizado")
        
        # Navegar para Relatórios Background
        print("\n📊 Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        
        # Aguardar tabela
        print("⏳ Aguardando carregamento da tabela...")
        page.wait_for_selector("tbody tr", timeout=60000)
        time.sleep(3)
        
        linhas = page.locator("tbody tr").all()
        print(f"📋 {len(linhas)} linhas encontradas na tabela")
        
        # =========================================================
        # RELATÓRIOS GERAIS (ELF e PDL)
        # =========================================================
        for nome_rel, conf in MAPEAMENTO_SP.items():
            if busca and not any(b.lower() in nome_rel.lower() for b in busca):
                continue
            
            print(f"\n🔍 Procurando: {nome_rel}")
            
            # Encontrar linha mais recente com status "Concluído"
            maior_dt = None
            linha_alvo = None
            for linha in linhas:
                texto = linha.inner_text()
                if nome_rel.lower() in texto.lower() and "Concluído" in texto:
                    m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
                    if m:
                        dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                        if maior_dt is None or dt > maior_dt:
                            maior_dt = dt
                            linha_alvo = linha
                            print(f"   📝 Encontrado: {texto[:80]}...")
            
            if linha_alvo:
                print(f"   📥 Baixando... (Data: {maior_dt})")
                with page.expect_download() as download_info:
                    linha_alvo.locator("button, a, img").first.click()
                
                download = download_info.value
                temp_path = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
                download.save_as(temp_path)
                
                with open(temp_path, 'rb') as f:
                    conteudo_zip = f.read()
                
                os.remove(temp_path)
                
                # Define o destino (local ou SharePoint)
                if is_github_actions():
                    destino = conf["pasta_sp"]
                else:
                    # Mapeia para pasta local
                    if "01_ELF_Diario" in conf["pasta_sp"]:
                        destino = PASTAS_LOCAIS["01_ELF_Diario"]
                    else:
                        destino = PASTAS_LOCAIS["02_PDL_Analitico"]
                
                print(f"   📤 Processando...")
                processar_zip(conteudo_zip, conf["nome_base"], conf["regra"], destino)
            else:
                print(f"   ⚠️ Nenhum relatório 'Concluído' encontrado")
        
        # =========================================================
        # INSTALAÇÕES NÃO VISITADAS (Diário e Histórico)
        # =========================================================
        print(f"\n🔍 Procurando: Instalações Não Visitadas")
        
        cand_04 = []  # Diário
        cand_05 = []  # Histórico
        
        for linha in linhas:
            texto = linha.inner_text()
            if "Instalações Não Visitadas" in texto and "Concluído" in texto:
                m_data = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
                m_params = re.findall(r"Dt(?:Ini|Fim)PrevLeitura=(\d{2}/\d{2}/\d{4})", texto)
                
                if m_data and len(m_params) >= 2:
                    dt_inclusao = datetime.strptime(m_data.group(1), "%d/%m/%Y %H:%M:%S")
                    # Se datas início e fim são diferentes = Diário (04)
                    if m_params[0] != m_params[1]:
                        cand_04.append({"dt": dt_inclusao, "linha": linha})
                        print(f"   📝 Diário encontrado: {m_params[0]} ≠ {m_params[1]}")
                    else:
                        cand_05.append({"dt": dt_inclusao, "linha": linha})
                        print(f"   📝 Histórico encontrado: {m_params[0]} = {m_params[1]}")
        
        # Download Diário (04)
        if cand_04:
            cand_04.sort(key=lambda x: x["dt"], reverse=True)
            alvo = cand_04[0]
            print(f"\n📥 Baixando Diário (04): {alvo['dt']}")
            with page.expect_download() as dw:
                alvo["linha"].locator("button, a, img").first.click()
            
            download = dw.value
            temp_path = f"temp_diario_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            download.save_as(temp_path)
            
            with open(temp_path, 'rb') as f:
                conteudo_zip = f.read()
            
            os.remove(temp_path)
            
            # Define destino
            if is_github_actions():
                destino = "BI_LEC/04_N_Visitado_Diario"
            else:
                destino = PASTAS_LOCAIS["04_N_Visitado_Diario"]
            
            print(f"   📤 Processando...")
            processar_zip(conteudo_zip, "Instalações Não Visitadas", "data_dia", destino)
        else:
            print(f"\n⚠️ Nenhum relatório Diário (04) encontrado")
        
        # Download Histórico (05)
        if cand_05:
            cand_05.sort(key=lambda x: x["dt"], reverse=True)
            alvo = cand_05[0]
            print(f"\n📥 Baixando Histórico (05): {alvo['dt']}")
            with page.expect_download() as dw:
                alvo["linha"].locator("button, a, img").first.click()
            
            download = dw.value
            temp_path = f"temp_historico_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            download.save_as(temp_path)
            
            with open(temp_path, 'rb') as f:
                conteudo_zip = f.read()
            
            os.remove(temp_path)
            
            # Define destino
            if is_github_actions():
                destino = "BI_LEC/05_N_Visitado_Historico"
            else:
                destino = PASTAS_LOCAIS["05_N_Visitado_Historico"]
            
            print(f"   📤 Processando...")
            processar_zip(conteudo_zip, "Instalações Não Visitadas", "data_dia", destino)
        else:
            print(f"⚠️ Nenhum relatório Histórico (05) encontrado")
        
        print("\n" + "="*60)
        print("✅ COLETOR FINALIZADO COM SUCESSO!")
        print("="*60)
        
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
        
        try:
            screenshot_path = f"erro_coletor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            page.screenshot(path=screenshot_path, full_page=True)
            print(f"📸 Screenshot salvo: {screenshot_path}")
        except:
            pass
        raise
    finally:
        context.close()
        
        # Limpeza no GitHub Actions
        if is_github_actions():
            user_dir = os.path.join(os.getcwd(), "temp_navegador")
            if os.path.exists(user_dir):
                import shutil
                shutil.rmtree(user_dir)
                print("🗑️ Diretório temporário removido")

# =================================================================
# 8. PONTO DE ENTRADA
# =================================================================
if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    print(f"\n🚀 Iniciando Coletor - Args: {args}")
    print(f"⏰ Horário: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    with sync_playwright() as playwright:
        run(playwright, args)