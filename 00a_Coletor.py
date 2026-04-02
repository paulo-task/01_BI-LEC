import os
import time
import re
import sys
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Tenta carregar .pass apenas se existir (ambiente local)
if os.path.exists(".pass"):
    load_dotenv(dotenv_path=".pass")
    print("✅ Arquivo .pass carregado")
else:
    print("ℹ️ Arquivo .pass não encontrado, usando variáveis de ambiente")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# VALIDAÇÃO CRÍTICA: Verifica se as credenciais foram carregadas
if not usuario or not senha:
    print("❌ ERRO FATAL: Credenciais não encontradas!")
    print(f"CPFL_USER: {'✅ configurado' if usuario else '❌ NÃO encontrado'}")
    print(f"CPFL_PASS: {'✅ configurado' if senha else '❌ NÃO encontrado'}")
    print("Verifique se as Secrets estão configuradas no GitHub:")
    print("  - Settings -> Secrets and variables -> Actions")
    print("  - CPFL_USER e CPFL_PASS devem estar cadastrados")
    sys.exit(1)  # Sai com erro para o GitHub Actions detectar

print("✅ Credenciais carregadas com sucesso!")

# Detecta se está no GitHub Actions
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

# Define se deve usar modo headless
if is_github_actions():
    USE_HEADLESS = True  # Força headless no GitHub
    print("🏗️ Executando no GitHub Actions - Modo headless ativado")
else:
    USE_HEADLESS = platform.system() != "Windows"
    print(f"💻 Executando localmente - Modo headless: {USE_HEADLESS}")

# --- MAPEAMENTO GERAL ---
if is_github_actions():
    BASE_DIR = os.getcwd()
    # Cria todas as pastas necessárias
    for pasta in ["01_ELF_Diario", "02_PDL_Analitico", "03_N_Lib_Fat_Diario", 
                  "04_N_Visitado_Diario", "05_N_Visitado_Historico", 
                  "06_Impedimentos", "07_Entregas", "01_LEC"]:
        os.makedirs(os.path.join(BASE_DIR, pasta), exist_ok=True)
    
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            os.path.join(BASE_DIR, "01_ELF_Diario"),
            "EFL",
            "data_dia_unico",
        ],
        "Produtividade Diária Leiturista - Analítico": [
            os.path.join(BASE_DIR, "02_PDL_Analitico"),
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
        ],
        "Inst. Não Liberadas Faturamento": [
            os.path.join(BASE_DIR, "03_N_Lib_Fat_Diario"),
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
        ],
        "Lista Impedimentos Aplicados": [
            os.path.join(BASE_DIR, "06_Impedimentos"),
            "Lista Impedimentos Aplicados",
            "data_mes",
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            os.path.join(BASE_DIR, "07_Entregas"),
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
        ],
    }
    PASTA_04 = os.path.join(BASE_DIR, "04_N_Visitado_Diario")
    PASTA_05 = os.path.join(BASE_DIR, "05_N_Visitado_Historico")
else:
    MAPEAMENTO = {
        "Efetividade de Leitura Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
            "EFL",
            "data_dia_unico",
        ],
        "Produtividade Diária Leiturista - Analítico": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
            "Produtividade Diária Leiturista - Analítico",
            "data_dia",
        ],
        "Inst. Não Liberadas Faturamento": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario",
            "Inst. Não Liberadas Faturamento",
            "substituir_puro",
        ],
        "Lista Impedimentos Aplicados": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos",
            "Lista Impedimentos Aplicados",
            "data_mes",
        ],
        "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [
            r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas",
            "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
            "data_mes",
        ],
    }
    PASTA_04 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario"
    PASTA_05 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico"

def tratar_arquivo(caminho_zip, pasta_destino, nome_base, regra):
    """Extrai o ZIP e aplica a regra de arquivo único por dia."""
    time.sleep(3)
    try:
        if not zipfile.is_zipfile(caminho_zip):
            print(f"❌ ZIP inválido: {caminho_zip}")
            return False

        os.makedirs(pasta_destino, exist_ok=True)

        with zipfile.ZipFile(caminho_zip, "r") as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()

            if regra == "data_dia_unico":
                prefixo_dia = f"{nome_base} - {agora.strftime('%Y_%m_%d')}"
                for f in os.listdir(pasta_destino):
                    if f.startswith(prefixo_dia) and f.endswith(ext):
                        try:
                            os.remove(os.path.join(pasta_destino, f))
                            print(f"🗑️ Removendo versão anterior: {f}")
                        except:
                            pass
                nome_final = f"{prefixo_dia}{ext}"
            elif regra == "data_dia":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m_%d')}{ext}"
            elif regra == "data_mes":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m')}{ext}"
            else:
                nome_final = f"{nome_base}{ext}"

            caminho_novo = os.path.join(pasta_destino, nome_final)
            z.extract(arquivo_interno, pasta_destino)
            caminho_extraido = os.path.join(pasta_destino, arquivo_interno)

            if os.path.abspath(caminho_extraido).lower() != os.path.abspath(caminho_novo).lower():
                if os.path.exists(caminho_novo):
                    os.remove(caminho_novo)
                os.rename(caminho_extraido, caminho_novo)

            print(f"✅ Sucesso: '{nome_final}' salvo em {pasta_destino}")
            return True

        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
        return True
    except Exception as e:
        print(f"❌ Erro no tratamento: {e}")
        return False

def run(playwright: Playwright, busca=None) -> None:
    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
        os.makedirs(user_dir, exist_ok=True)
    else:
        user_dir = r"C:\Temp\dados_navegador_coletor"

    print(f"🔍 Modo headless: {USE_HEADLESS}")
    print(f"📁 Diretório do usuário: {user_dir}")
    print(f"🐍 Python: {platform.python_version()}")
    print(f"💻 Sistema: {platform.system()}")

    # Argumentos essenciais para Linux (GitHub Actions)
    launch_args = ['--no-sandbox', '--disable-setuid-sandbox']
    if is_github_actions():
        launch_args.append('--disable-dev-shm-usage')
        launch_args.append('--disable-gpu')

    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,  # ← CRÍTICO para Linux!
        no_viewport=True,
        slow_mo=1000,  # Mais lento para garantir
    )
    page = context.pages[0]

    # Timeout maior no GitHub (10 minutos)
    timeout_ms = 600000 if is_github_actions() else 300000
    page.set_default_timeout(timeout_ms)

    try:
        print("🌐 Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)
        
        # Aguarda a página carregar
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

        # Aguarda a tabela carregar
        print("⏳ Aguardando tabela carregar...")
        page.wait_for_selector("tbody tr", state="visible", timeout=300000)
        time.sleep(5)  # Reduzido de 30 para 5 segundos

        linhas = page.locator("tbody tr").all()
        print(f"📋 Encontradas {len(linhas)} linhas na tabela")

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
                print(f"📥 Baixando: {nome_rel} (Data: {maior_dt})")
                os.makedirs(conf[0], exist_ok=True)
                with page.expect_download(timeout=600000) as dw_info:
                    linha_alvo.locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir|Remover", re.I)
                    ).first.click()
                tmp_path = os.path.join(conf[0], f"temp_{conf[1][:5]}.zip")
                dw_info.value.save_as(tmp_path)
                tratar_arquivo(tmp_path, conf[0], conf[1], conf[2])
            else:
                print(f"⚠️ Nenhum relatório concluído para: {nome_rel}")

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
                with page.expect_download(timeout=600000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_04, "temp_04.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_04, "Instalações Não Visitadas", "data_dia")

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
                tratar_arquivo(path_zip, PASTA_05, "Instalações Não Visitadas", "data_dia")

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
    with sync_playwright() as playwright:
        run(playwright, args)