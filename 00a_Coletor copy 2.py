import os
import time
import re
import sys
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .pass
try:
    load_dotenv(dotenv_path=".pass")
    print("✅ Arquivo .pass carregado")
except:
    print("ℹ️ Arquivo .pass não encontrado, usando variáveis de ambiente")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# Detecta se está no GitHub Actions
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

# Define se deve usar modo headless
if is_github_actions():
    USE_HEADLESS = True
    print("🏗️ Modo headless: GitHub Actions")
else:
    USE_HEADLESS = False
    print("🪟 Modo visível: Windows Local")

# Configuração dos caminhos conforme o ambiente
if is_github_actions():
    # No GitHub: pastas temporárias (só para processamento)
    BASE_DIR = os.getcwd()
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            os.path.join(BASE_DIR, "01_ELF_Diario"),
            "EFL",
            "data_dia_unico",
            "BI_LEC/01_ELF_Diario"
        ],
        "Produtividade Diária Leiturista - Analítico": [
            os.path.join(BASE_DIR, "02_PDL_Analitico"),
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
            "BI_LEC/02_PDL_Analitico"
        ],
        "Inst. Não Liberadas Faturamento": [
            os.path.join(BASE_DIR, "03_N_Lib_Fat_Diario"),
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
            "BI_LEC/03_N_Lib_Fat_Diario"
        ],
        "Lista Impedimentos Aplicados": [
            os.path.join(BASE_DIR, "06_Impedimentos"),
            "Lista Impedimentos Aplicados",
            "data_mes",
            "BI_LEC/06_Impedimentos"
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            os.path.join(BASE_DIR, "07_Entregas"),
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
            "BI_LEC/07_Entregas"
        ],
    }
    PASTA_04 = os.path.join(BASE_DIR, "04_N_Visitado_Diario")
    PASTA_05 = os.path.join(BASE_DIR, "05_N_Visitado_Historico")
    PASTA_04_SP = "BI_LEC/04_N_Visitado_Diario"
    PASTA_05_SP = "BI_LEC/05_N_Visitado_Historico"
else:
    # No Windows: caminhos originais do seu PC
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
            "EFL",
            "data_dia_unico",
            "BI_LEC/01_ELF_Diario"
        ],
        "Produtividade Diária Leiturista - Analítico": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
            "BI_LEC/02_PDL_Analitico"
        ],
        "Inst. Não Liberadas Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario",
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
            "BI_LEC/03_N_Lib_Fat_Diario"
        ],
        "Lista Impedimentos Aplicados": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos",
            "Lista Impedimentos Aplicados",
            "data_mes",
            "BI_LEC/06_Impedimentos"
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas",
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
            "BI_LEC/07_Entregas"
        ],
    }
    PASTA_04 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario"
    PASTA_05 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico"
    PASTA_04_SP = "BI_LEC/04_N_Visitado_Diario"
    PASTA_05_SP = "BI_LEC/05_N_Visitado_Historico"

def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo diretamente para o SharePoint (apenas no GitHub)"""
    if not is_github_actions():
        print("📁 Modo local: não enviando para SharePoint")
        return True
    
    print(f"🔐 Iniciando upload para SharePoint: {nome_arquivo}")
    print(f"📁 Pasta SharePoint: {pasta_sharepoint}")
    
    try:
        import requests
        from urllib.parse import urlparse
        
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID")
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
        SP_TENANT_ID = os.getenv("SP_TENANT_ID")
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        # Verifica se as credenciais existem
        if not SP_CLIENT_ID:
            print("❌ SP_CLIENT_ID não encontrado nas variáveis de ambiente!")
            return False
        if not SP_CLIENT_SECRET:
            print("❌ SP_CLIENT_SECRET não encontrado!")
            return False
        if not SP_TENANT_ID:
            print("❌ SP_TENANT_ID não encontrado!")
            return False
            
        print(f"✅ Credenciais encontradas: CLIENT_ID={SP_CLIENT_ID[:10]}...")
        
        # 1. Obter token
        print("📡 Obtendo token de acesso...")
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
            print(f"Resposta: {r.text[:200]}")
            return False
            
        token = r.json()['access_token']
        print("✅ Token obtido com sucesso")
        
        # 2. Obter Site ID
        print("📍 Obtendo Site ID...")
        parsed = urlparse(SITE_URL)
        host = parsed.netloc
        site_path = parsed.path.strip("/")
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url_site, headers=headers)
        
        if r.status_code != 200:
            print(f"❌ Erro ao obter Site ID: {r.status_code}")
            print(f"Resposta: {r.text[:200]}")
            return False
            
        site_id = r.json()["id"]
        print(f"✅ Site ID: {site_id[:50]}...")
        
        # 3. Obter Drive ID da biblioteca Workspace
        print("🔍 Buscando biblioteca Workspace...")
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        r = requests.get(url_drives, headers={"Authorization": f"Bearer {token}"})
        
        if r.status_code != 200:
            print(f"❌ Erro ao listar drives: {r.status_code}")
            return False
            
        drive_id = None
        drives = r.json().get('value', [])
        print(f"📁 Bibliotecas encontradas: {len(drives)}")
        for drive in drives:
            print(f"   - {drive.get('name')}")
            if drive.get('name') == 'Workspace':
                drive_id = drive.get('id')
                print(f"✅ Biblioteca Workspace encontrada! ID: {drive_id[:30]}...")
                break
        
        if not drive_id:
            print("❌ Biblioteca 'Workspace' não encontrada!")
            return False
        
        # 4. Upload do arquivo
        print(f"📤 Enviando arquivo: {pasta_sharepoint}/{nome_arquivo}")
        url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content"
        headers_upload = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        r = requests.put(url_upload, headers=headers_upload, data=conteudo_bytes)
        
        if r.status_code in [200, 201]:
            print(f"✅ Upload realizado com sucesso: {nome_arquivo}")
            web_url = r.json().get('webUrl')
            print(f"🔗 URL: {web_url}")
            return True
        else:
            print(f"❌ Erro no upload: Status {r.status_code}")
            print(f"Resposta: {r.text[:200]}")
            return False
            
    except Exception as e:
        print(f"❌ Exceção no upload SharePoint: {e}")
        import traceback
        traceback.print_exc()
        return False

def tratar_arquivo(caminho_zip, pasta_destino, nome_base, regra, pasta_sharepoint):
    """Extrai o ZIP e salva (localmente no PC ou SharePoint no GitHub)"""
    time.sleep(3)
    try:
        if not zipfile.is_zipfile(caminho_zip):
            print(f"Erro: ZIP inválido: {caminho_zip}")
            return

        with zipfile.ZipFile(caminho_zip, "r") as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()

            if regra == "data_dia_unico":
                prefixo_dia = f"{nome_base} - {agora.strftime('%Y_%m_%d')}"
                nome_final = f"{prefixo_dia}{ext}"
            elif regra == "data_dia":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m_%d')}{ext}"
            elif regra == "data_mes":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m')}{ext}"
            else:
                nome_final = f"{nome_base}{ext}"

            # Lê o conteúdo do arquivo
            with z.open(arquivo_interno) as arquivo_zip:
                conteudo = arquivo_zip.read()

            if is_github_actions():
                # No GitHub: envia para o SharePoint
                print(f"📤 [GITHUB] Enviando para SharePoint: {nome_final} -> {pasta_sharepoint}")
                sucesso = upload_to_sharepoint(conteudo, nome_final, pasta_sharepoint)
                if sucesso:
                    print(f"✅ Upload concluído: {nome_final}")
                else:
                    print(f"❌ Falha no upload: {nome_final}")
            else:
                # No Windows: salva na pasta local
                print(f"💾 [LOCAL] Salvando em: {pasta_destino}")
                os.makedirs(pasta_destino, exist_ok=True)
                
                # Remove versão anterior se existir (para ELF)
                if regra == "data_dia_unico":
                    for f in os.listdir(pasta_destino):
                        if f.startswith(prefixo_dia) and f.endswith(ext):
                            try:
                                os.remove(os.path.join(pasta_destino, f))
                                print(f"Limpando versão anterior: {f}")
                            except:
                                pass
                
                caminho_novo = os.path.join(pasta_destino, nome_final)
                with open(caminho_novo, 'wb') as f:
                    f.write(conteudo)
                print(f"✔ Sucesso: '{nome_final}' salvo em {pasta_destino}")

        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
    except Exception as e:
        print(f"Erro no tratamento: {e}")

def run(playwright: Playwright, busca=None) -> None:
    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
        os.makedirs(user_dir, exist_ok=True)
    else:
        user_dir = r"C:\Temp\dados_navegador_coletor"

    print(f"🔍 Modo headless: {USE_HEADLESS}")
    print(f"📁 Diretório do usuário: {user_dir}")
    print(f"🏗️ Ambiente: {'GitHub Actions' if is_github_actions() else 'Windows Local'}")

    # Argumentos para Linux (GitHub)
    launch_args = ['--no-sandbox', '--disable-setuid-sandbox']
    if is_github_actions():
        launch_args.append('--disable-dev-shm-usage')

    # Inicia o navegador
    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,
        no_viewport=True,
        slow_mo=1000,
    )
    page = context.pages[0]

    # Timeout global: 10 minutos (600.000ms)
    page.set_default_timeout(600000)

    try:
        print("🌐 Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)

        page.wait_for_load_state("networkidle", timeout=30000)

        # Login
        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=30000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()
            print("✅ Login realizado")
            page.wait_for_load_state("networkidle", timeout=30000)

        print("📊 Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_load_state("networkidle", timeout=30000)

        # Espera a tabela carregar
        print("⏳ Aguardando tabela carregar...")
        page.wait_for_selector("tbody tr", state="visible", timeout=300000)
        print("✅ Tabela carregada!")
        
        # Aguarda renderização completa
        print("Aguardando 30s para renderização total da tabela...")
        time.sleep(30)

        linhas = page.locator("tbody tr").all()
        print(f"📋 Encontradas {len(linhas)} linhas na tabela")
        
        if len(linhas) > 0:
            print(f"📝 Exemplo da primeira linha: {linhas[0].inner_text()[:200]}")

        # --- RELATÓRIOS GERAIS ---
        for nome_rel, conf in MAPEAMENTO.items():
            if busca and not any(t.lower() in nome_rel.lower() for t in busca):
                continue

            print(f"🔎 Procurando: {nome_rel}")
            maior_dt = None
            linha_alvo = None
            for linha in linhas:
                t = linha.inner_text()
                if nome_rel.lower() in t.lower() and "Concluído" in t:
                    m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    if m:
                        dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                        if maior_dt is None or dt > maior_dt:
                            maior_dt = dt
                            linha_alvo = linha

            if linha_alvo:
                print(f"📥 Baixando mais recente: {nome_rel} (Data: {maior_dt})")
                os.makedirs(conf[0], exist_ok=True)
                with page.expect_download(timeout=600000) as dw_info:
                    linha_alvo.locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir|Remover", re.I)
                    ).first.click()
                tmp_path = os.path.join(conf[0], f"temp_{conf[1][:5]}.zip")
                dw_info.value.save_as(tmp_path)
                tratar_arquivo(tmp_path, conf[0], conf[1], conf[2], conf[3])
            else:
                print(f"⚠️ Nenhum relatório 'Concluído' encontrado para: {nome_rel}")

        # --- NÃO VISITADAS (04 e 05) ---
        if not busca or any("visitada" in b.lower() for b in busca):
            cand_04 = []
            cand_05 = []
            for linha in linhas:
                t = linha.inner_text()
                if "Instalações Não Visitadas" in t and "Concluído" in t:
                    m_data = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    m_params = re.findall(r"Dt(?:Ini|Fim)PrevLeitura=(\d{2}/\d{2}/\d{4})", t)
                    if m_data and len(m_params) >= 2:
                        dt_inclusao = datetime.strptime(m_data.group(1), "%d/%m/%Y %H:%M:%S")
                        if m_params[0] != m_params[1]:
                            cand_04.append({"dt": dt_inclusao, "linha": linha})
                        else:
                            cand_05.append({"dt": dt_inclusao, "linha": linha})

            # Download Diário 04
            if cand_04:
                cand_04.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_04[0]
                print(f"📥 Baixando Diário (04): {alvo['dt']}")
                os.makedirs(PASTA_04, exist_ok=True)
                with page.expect_download(timeout=600000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_04, "temp_04.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_04, "Instalações Não Visitadas", "data_dia", PASTA_04_SP)
            else:
                print("⚠️ Nenhum relatório 'Instalações Não Visitadas' (Diário) encontrado")

            # Download Histórico 05
            if cand_05:
                cand_05.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_05[0]
                print(f"📥 Baixando Histórico (05): {alvo['dt']}")
                os.makedirs(PASTA_05, exist_ok=True)
                with page.expect_download(timeout=600000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_05, "temp_05.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_05, "Instalações Não Visitadas", "data_dia", PASTA_05_SP)
            else:
                print("⚠️ Nenhum relatório 'Instalações Não Visitadas' (Histórico) encontrado")

        print("✅ Coletor finalizado com sucesso!")

    except Exception as e:
        print(f"❌ Erro Crítico: {e}")
        import traceback
        traceback.print_exc()
        
        try:
            screenshot_path = f"erro_coletor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            page.screenshot(path=screenshot_path, full_page=True)
            print(f"📸 Screenshot salvo: {screenshot_path}")
        except:
            print("⚠️ Não foi possível salvar screenshot")
        raise
    finally:
        time.sleep(5)
        context.close()

if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    print(f"🚀 Iniciando Coletor - Args: {args}")
    print(f"⏰ Horário de início: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with sync_playwright() as playwright:
        run(playwright, args)