import os
import time
import re
import sys
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# Detecta se está no GitHub Actions
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"


# Define se deve usar modo headless
USE_HEADLESS = is_github_actions() or platform.system() != "Windows"

# --- MAPEAMENTO GERAL - CORRIGIDO PARA AMBIENTE LINUX ---
# No GitHub, vamos usar caminhos relativos dentro do repositório
if is_github_actions():
    # Caminhos relativos para o GitHub Actions
    BASE_DIR = os.getcwd()  # Diretório atual do repositório
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
    # Caminhos originais do Windows
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
    """Extrai o ZIP e aplica a regra de arquivo único por dia para o ELF."""
    time.sleep(3)
    try:
        if not zipfile.is_zipfile(caminho_zip):
            print(f"Erro: ZIP inválido: {caminho_zip}")
            return

        # Cria a pasta de destino se não existir
        os.makedirs(pasta_destino, exist_ok=True)

        with zipfile.ZipFile(caminho_zip, "r") as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()

            # --- NOVA REGRA PARA ELF: UM ARQUIVO POR DIA ---
            if regra == "data_dia_unico":
                prefixo_dia = f"{nome_base} - {agora.strftime('%Y_%m_%d')}"
                # Procura e remove arquivos do mesmo dia para não acumular
                for f in os.listdir(pasta_destino):
                    if f.startswith(prefixo_dia) and f.endswith(ext):
                        try:
                            os.remove(os.path.join(pasta_destino, f))
                            print(f"Limpando versão anterior do dia: {f}")
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

            if (
                os.path.abspath(caminho_extraido).lower()
                != os.path.abspath(caminho_novo).lower()
            ):
                if os.path.exists(caminho_novo):
                    os.remove(caminho_novo)
                os.rename(caminho_extraido, caminho_novo)

            print(f"✔ Sucesso: '{nome_final}' salvo.")

        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
    except Exception as e:
        print(f"Erro no tratamento: {e}")


def run(playwright: Playwright, busca=None) -> None:
    # Usa um diretório temporário no GitHub ou o original no Windows
    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
    else:
        user_dir = r"C:\Temp\dados_navegador_coletor"

    print(f"🔍 Modo headless: {USE_HEADLESS}")
    print(f"📁 Diretório do usuário: {user_dir}")

    # Inicia o navegador com headless configurado automaticamente
    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,  # ← CORREÇÃO PRINCIPAL
        no_viewport=True,
        slow_mo=500,
    )
    page = context.pages[0]

    # Timeout global aumentado para 5 minutos (300.000ms)
    page.set_default_timeout(300000)

    try:
        print("Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")

        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=15000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()

        print("Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()

        # Espera a tabela carregar
        page.wait_for_selector("tbody tr", state="visible", timeout=180000)
        print("Aguardando 30s para renderização total da tabela...")
        time.sleep(30)

        linhas = page.locator("tbody tr").all()

        # --- PARTE 1: RELATÓRIOS GERAIS (INCLUINDO ELF) ---
        for nome_rel, conf in MAPEAMENTO.items():
            if busca and not any(t.lower() in nome_rel.lower() for t in busca):
                continue

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
                print(f"Baixando mais recente: {nome_rel}")
                os.makedirs(conf[0], exist_ok=True)
                with page.expect_download(timeout=300000) as dw_info:
                    linha_alvo.locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir|Remover", re.I)
                    ).first.click()
                tmp_path = os.path.join(conf[0], f"temp_{conf[1][:5]}.zip")
                dw_info.value.save_as(tmp_path)
                tratar_arquivo(tmp_path, conf[0], conf[1], conf[2])

        # --- PARTE 2: NÃO VISITADAS (04 e 05) ---
        if not busca or any("visitada" in b.lower() for b in busca):
            cand_04 = []
            cand_05 = []
            for linha in linhas:
                t = linha.inner_text()
                if "Instalações Não Visitadas" in t and "Concluído" in t:
                    m_data = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    m_params = re.findall(
                        r"Dt(?:Ini|Fim)PrevLeitura=(\d{2}/\d{2}/\d{4})", t
                    )
                    if m_data and len(m_params) >= 2:
                        dt_inclusao = datetime.strptime(
                            m_data.group(1), "%d/%m/%Y %H:%M:%S"
                        )
                        if m_params[0] != m_params[1]:
                            cand_04.append({"dt": dt_inclusao, "linha": linha})
                        else:
                            cand_05.append({"dt": dt_inclusao, "linha": linha})

            # Download Diário 04
            if cand_04:
                cand_04.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_04[0]
                print(f"Baixando Diário (04): {alvo['dt']}")
                os.makedirs(PASTA_04, exist_ok=True)
                with page.expect_download(timeout=300000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_04, "temp_04.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(
                    path_zip, PASTA_04, "Instalações Não Visitadas", "data_dia"
                )

            # Download Histórico 05
            if cand_05:
                cand_05.sort(key=lambda x: x["dt"], reverse=True)
                alvo = cand_05[0]
                print(f"Baixando Histórico (05): {alvo['dt']}")
                os.makedirs(PASTA_05, exist_ok=True)
                with page.expect_download(timeout=300000) as dw:
                    alvo["linha"].locator("button, a, img").filter(
                        has_not_text=re.compile(r"Excluir", re.I)
                    ).first.click()
                path_zip = os.path.join(PASTA_05, "temp_05.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(
                    path_zip, PASTA_05, "Instalações Não Visitadas", "data_dia"
                )

    except Exception as e:
        print(f"Erro Crítico: {e}")
        # Tira screenshot em caso de erro (útil para debug)
        try:
            page.screenshot(path="erro_coletor.png")
            print("📸 Screenshot salvo: erro_coletor.png")
        except:
            pass
        raise
    finally:
        time.sleep(5)
        context.close()


if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    with sync_playwright() as playwright:
        run(playwright, args)
