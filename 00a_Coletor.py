import os
import time
import re
import sys
import zipfile
import platform
import io
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .pass
if os.path.exists(".pass"):
    load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# Detecta se está no GitHub Actions
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

# Define se deve usar modo headless
USE_HEADLESS = is_github_actions() or platform.system() != "Windows"

# Configuração do SharePoint (apenas para GitHub)
SP_CLIENT_ID = os.getenv("SP_CLIENT_ID")
SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
SP_SITE_URL = os.getenv("SP_SITE_URL")

# --- MAPEAMENTO DOS CAMINHOS ---
if is_github_actions():
    # No GitHub: pastas temporárias (só para processamento)
    BASE_DIR = os.getcwd()
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            os.path.join(BASE_DIR, "01_ELF_Diario"),
            "EFL",
            "data_dia_unico",
            "01_ELF_Diario"  # nome da pasta no SharePoint
        ],
        "Produtividade Diária Leiturista - Analítico": [
            os.path.join(BASE_DIR, "02_PDL_Analitico"),
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
            "02_PDL_Analitico"
        ],
        "Inst. Não Liberadas Faturamento": [
            os.path.join(BASE_DIR, "03_N_Lib_Fat_Diario"),
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
            "03_N_Lib_Fat_Diario"
        ],
        "Lista Impedimentos Aplicados": [
            os.path.join(BASE_DIR, "06_Impedimentos"),
            "Lista Impedimentos Aplicados",
            "data_mes",
            "06_Impedimentos"
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            os.path.join(BASE_DIR, "07_Entregas"),
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
            "07_Entregas"
        ],
    }
    PASTA_04 = os.path.join(BASE_DIR, "04_N_Visitado_Diario")
    PASTA_05 = os.path.join(BASE_DIR, "05_N_Visitado_Historico")
    PASTA_04_SP = "04_N_Visitado_Diario"
    PASTA_05_SP = "05_N_Visitado_Historico"
else:
    # No Windows: caminhos originais do seu PC
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
            "EFL",
            "data_dia_unico",
            "01_ELF_Diario"
        ],
        "Produtividade Diária Leiturista - Analítico": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
            "02_PDL_Analitico"
        ],
        "Inst. Não Liberadas Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario",
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
            "03_N_Lib_Fat_Diario"
        ],
        "Lista Impedimentos Aplicados": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos",
            "Lista Impedimentos Aplicados",
            "data_mes",
            "06_Impedimentos"
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas",
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
            "07_Entregas"
        ],
    }
    PASTA_04 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario"
    PASTA_05 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico"
    PASTA_04_SP = "04_N_Visitado_Diario"
    PASTA_05_SP = "05_N_Visitado_Historico"

def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo diretamente para o SharePoint (apenas no GitHub)"""
    if not is_github_actions():
        return True
    
    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.client_credential import ClientCredential
        
        print(f"📤 Enviando para SharePoint: {nome_arquivo} -> {pasta_sharepoint}")
        
        credentials = ClientCredential(SP_CLIENT_ID, SP_CLIENT_SECRET)
        ctx = ClientContext(SP_SITE_URL).with_credentials(credentials)
        
        caminho_sharepoint = f"Documentos Compartilhados/BI_LEC/{pasta_sharepoint}"
        
        # Garante que a pasta existe
        folder = ctx.web.get_folder_by_server_relative_url(caminho_sharepoint)
        ctx.load(folder)
        ctx.execute_query()
        
        # Upload do arquivo
        arquivo_stream = io.BytesIO(conteudo_bytes)
        target_file = folder.upload_file(nome_arquivo, arquivo_stream.read())
        ctx.execute_query()
        
        print(f"✅ Upload concluído: {nome_arquivo}")
        return True
    except Exception as e:
        print(f"❌ Erro no upload para SharePoint: {e}")
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
                upload_to_sharepoint(conteudo, nome_final, pasta_sharepoint)
            else:
                # No Windows: salva na pasta local
                os.makedirs(pasta_destino, exist_ok=True)
                
                # Remove versão anterior se existir
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

    # Argumentos para Linux
    launch_args = ['--no-sandbox', '--disable-setuid-sandbox']
    if is_github_actions():
        launch_args.append('--disable-dev-shm-usage')

    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,
        no_viewport=True,
        slow_mo=500,
    )
    page = context.pages[0]
    page.set_default_timeout(300000)

    try:
        print("🌐 Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")

        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=15000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()

        print("📊 Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_selector("tbody tr", state="visible", timeout=180000)
        print("Aguardando renderização da tabela...")
        time.sleep(5)

        linhas = page.locator("tbody tr").all()
        print(f"📋 Encontradas {len(linhas)} linhas")

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
                print(f"📥 Baixando: {nome_rel}")
                os.makedirs(conf[0], exist_ok=True)
                with page.expect_download(timeout=300000) as dw_info:
                    linha_alvo.locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir|Remover", re.I)
                    ).first.click()
                tmp_path = os.path.join(conf[0], f"temp_{conf[1][:5]}.zip")
                dw_info.value.save_as(tmp_path)
                tratar_arquivo(tmp_path, conf[0], conf[1], conf[2], conf[3])

        # --- NÃO VISITADAS ---
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

            if cand_04:
                cand_04.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_04[0]
                print(f"📥 Baixando Diário (04): {alvo['dt']}")
                os.makedirs(PASTA_04, exist_ok=True)
                with page.expect_download(timeout=300000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_04, "temp_04.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_04, "Instalações Não Visitadas", "data_dia", PASTA_04_SP)

            if cand_05:
                cand_05.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_05[0]
                print(f"📥 Baixando Histórico (05): {alvo['dt']}")
                os.makedirs(PASTA_05, exist_ok=True)
                with page.expect_download(timeout=300000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_05, "temp_05.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_05, "Instalações Não Visitadas", "data_dia", PASTA_05_SP)

        print("✅ Coletor finalizado com sucesso!")

    except Exception as e:
        print(f"❌ Erro Crítico: {e}")
        try:
            page.screenshot(path="erro_coletor.png")
            print("📸 Screenshot salvo")
        except:
            pass
        raise
    finally:
        time.sleep(5)
        context.close()

if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    print(f"🚀 Iniciando Coletor - Args: {args}")
    with sync_playwright() as playwright:
        run(playwright, args)