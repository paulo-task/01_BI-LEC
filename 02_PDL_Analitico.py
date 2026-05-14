import os
import platform
from datetime import datetime
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


def run(playwright: Playwright) -> None:
    is_headless = platform.system() != "Windows"
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Produtividade Diária..."
    )

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

    browser = playwright.chromium.launch(headless=is_headless, slow_mo=500)
    context = browser.new_context(viewport={'width': 1920, 'height': 1080})
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
        page.get_by_text("Produtividade Diária").click()

        # 3. DATAS DINÂMICAS
        data_hoje = datetime.now().strftime("%d/%m/%Y")
        print(f"Preenchendo datas: {data_hoje}")

        page.locator('input[name="j_idt93:0:j_idt99"]').click()
        page.keyboard.type(data_hoje)
        page.keyboard.press("Tab")
        page.keyboard.type(data_hoje)
        page.keyboard.press("Tab")

        # 4. FILTROS

        # --- EMPRESAS ---
        print("Preenchendo Empresas...")
        for emp in ["PAULISTA", "PIRATININGA"]:
            page.keyboard.type(emp)
            page.keyboard.press("Enter")
            page.wait_for_timeout(200)
        page.keyboard.press("Tab")

        # --- REGIONAIS ---
        print("Preenchendo Regionais...")
        for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
            page.keyboard.type(reg)
            page.keyboard.press("Enter")
            page.wait_for_timeout(200)
        page.keyboard.press("Tab")

        # --- CIDADES ---
        print("Preenchendo Cidades...")
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
            page.keyboard.press("Enter")
            page.wait_for_timeout(100)

        # 5. GERAR RELATÓRIO
        print("Clicando em Gerar Background...")
        page.keyboard.press("Tab")
        page.get_by_role("button", name="Gerar Background").click()

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
        try:
            page.screenshot(path="erro_pdl_analitico.png", full_page=True)
            print("📸 Screenshot de erro salvo como erro_pdl_analitico.png")
        except:
            pass
        raise

    finally:
        context.close()
        browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
