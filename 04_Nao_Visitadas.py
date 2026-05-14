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
    data = datetime.now().replace(day=1)
    if data.weekday() == 5:  # Sábado
        data = data + timedelta(days=2)
    elif data.weekday() == 6:  # Domingo
        data = data + timedelta(days=1)
    return data.strftime("%d/%m/%Y")


def run(playwright: Playwright) -> None:
    is_headless = platform.system() != "Windows"
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Instalações Não Visitadas..."
    )

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

    browser = playwright.chromium.launch(headless=is_headless, slow_mo=500)
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
        page.get_by_text("Instalações Não Visitadas").click()

        # 3. FILTROS
        print("Preenchendo Filtros de Localidade...")
        page.locator(".selectize-input").first.click()

        for emp in ["PAULISTA", "PIRATININGA"]:
            page.keyboard.type(emp)
            page.get_by_text(emp, exact=True).first.click()
        page.keyboard.press("Tab")

        for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
            page.keyboard.type(reg)
            page.get_by_text(reg, exact=True).first.click()
        page.keyboard.press("Tab")

        cidades = [
            "BAURU",
            "BOTUCATU",
            "JAU",
            "MARILIA",
            "INDAIATUBA",
            "JUNDIAI [B]",
            "ITU",
            "MAIRINQUE",
            "SOROCABA [B]",
            "JUNDIAI [A]",
            "SALTO [A]",
            "SOROCABA [A]",
        ]
        for cid in cidades:
            page.keyboard.type(cid)
            try:
                page.get_by_text(cid, exact=True).first.click()
            except:
                page.keyboard.press("Enter")
        page.keyboard.press("Tab")

        # 4. DATAS DINÂMICAS (1º Útil até Hoje)
        data_inicio = get_primeiro_dia_util()
        data_hoje = datetime.now().strftime("%d/%m/%Y")
        print(f"Preenchendo: {data_inicio} até {data_hoje}")

        campo_data_ini = page.locator("input[name*='j_idt99']").first
        campo_data_ini.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(data_inicio)
        page.keyboard.press("Tab")

        campo_data_fim = page.locator("input[name*='j_idt99']").nth(1)
        campo_data_fim.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(data_hoje)
        page.keyboard.press("Tab")

        # 5. GERAR RELATÓRIO
        print("Gerando Relatório em Background...")
        page.get_by_role("button", name="Gerar Background").click()
        page.wait_for_timeout(5000)
        print(
            f"[{datetime.now().strftime('%H:%M:%S')}] Sucesso! O relatório está sendo gerado."
        )

        # Tentativa de Logout para liberar a sessão no servidor da CPFL
        try:
            print("Deslogando do sistema para liberar a sessão...")
            import re
            page.get_by_role("menuitem", name=re.compile("Logout", re.IGNORECASE)).click(force=True, timeout=5000)
            page.wait_for_load_state("networkidle", timeout=5000)
            print("Sessão encerrada com sucesso.")
        except Exception as e:
            print(f"Aviso: Não foi possível deslogar automaticamente. {e}")

    except Exception as e:
        print(f"❌ Erro durante a execução: {e}")
        raise

    finally:
        context.close()
        browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
