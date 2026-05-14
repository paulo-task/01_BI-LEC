import os
import platform
from datetime import datetime, timedelta
from playwright.sync_api import Playwright, sync_playwright

# 1. TRATAMENTO DE VARIÁVEIS (PC vs GITHUB)
try:
    from dotenv import load_dotenv

    base_dir = os.path.dirname(os.path.abspath(__file__))
    pass_file = os.path.join(base_dir, ".pass")
    if os.path.exists(pass_file):
        load_dotenv(dotenv_path=pass_file)
except ImportError:
    pass

usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")


def get_primeiro_dia_util():
    hoje = datetime.now()
    data = hoje.replace(day=1)
    if data.weekday() == 5:  # Sábado
        data = data + timedelta(days=2)
    elif data.weekday() == 6:  # Domingo
        data = data + timedelta(days=1)
    return data.strftime("%d/%m/%Y")


def run(playwright: Playwright) -> None:
    is_headless = platform.system() != "Windows"
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Inst. Não Liberadas..."
    )

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

    browser = playwright.chromium.launch(headless=is_headless, slow_mo=600)
    context = browser.new_context()
    page = context.new_page()

    try:
        # 1. ACESSO E LOGIN
        print("Acessando página de login...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)

        # Espera obrigatória de até 60s para a página estar pronta
        usuario_input = page.get_by_role("textbox", name="Usuário")
        usuario_input.wait_for(state="visible", timeout=60000)
        usuario_input.fill(usuario)
        
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()

        # Aguarda a tela pós-login carregar completamente
        page.wait_for_load_state("networkidle", timeout=60000)

        # 2. NAVEGAÇÃO
        print("Navegando no menu...")
        menu_relatorios = page.get_by_text("Relatórios").nth(1)
        menu_relatorios.wait_for(state="visible", timeout=60000)
        menu_relatorios.click()
        
        # Atraso estratégico para aguardar a abertura visual do dropdown do menu
        page.wait_for_timeout(2000)
        
        page.get_by_role("link", name="LEC", exact=True).click()
        page.get_by_text("Inst. Não Liberadas").click()

        # 3. DATAS DINÂMICAS
        data_inicio = get_primeiro_dia_util()
        data_fim = datetime.now().strftime("%d/%m/%Y")
        print(f"Data Início (1º útil): {data_inicio}")
        print(f"Data Fim (Hoje): {data_fim}")

        campo_data = page.locator('input[name="j_idt93:0:j_idt99"]')
        campo_data.click()
        page.keyboard.type(data_inicio)
        page.keyboard.press("Tab")
        page.keyboard.type(data_fim)
        page.keyboard.press("Tab")

        # 4. FILTROS
        filtros = {
            "Empresas": ["PAULISTA", "PIRATININGA"],
            "Regionais": ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"],
            "Cidades": [
                "BAURU",
                "BOTUCATU",
                "JAU",
                "MARILIA",
                "INDAIATUBA",
                "JUNDIAI [B]",
                "ITU",
                "MAIRINQUE",
                "SOROCABA [B]",
            ],
        }

        for categoria, valores in filtros.items():
            print(f"Preenchendo {categoria}...")
            for valor in valores:
                page.keyboard.type(valor)
                try:
                    page.get_by_text(valor, exact=True).first.click()
                except:
                    page.keyboard.press("Enter")
                page.wait_for_timeout(300)
            page.keyboard.press("Tab")

        # 5. CAMPO BT
        try:
            page.locator("select[id*='rel-parametro-27']").select_option("BT")
        except:
            page.get_by_label("Tensão").select_option("BT")

        # 6. GERAR RELATÓRIO
        print("Gerando relatório...")
        page.get_by_role("button", name="Gerar Background").click()
        page.wait_for_timeout(3000)
        print(
            f"[{datetime.now().strftime('%H:%M:%S')}] Sucesso! O relatório está sendo gerado."
        )

    except Exception as e:
        print(f"❌ Erro durante a execução: {e}")
        raise

    finally:
        context.close()
        browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
