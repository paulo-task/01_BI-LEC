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
    # No PC mostra janela, no GitHub/Codespace roda invisível
    is_headless = platform.system() != "Windows"

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando automação (Headless: {is_headless})...")

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

    # slow_mo=500 igual ao PC — controla o ritmo sem precisar de waits extras
    browser = playwright.chromium.launch(headless=is_headless, slow_mo=500)
    context = browser.new_context()
    page = context.new_page()

    try:
        # 1. ACESSO E LOGIN
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        page.get_by_role("textbox", name="Usuário").fill(usuario)
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()

        page.wait_for_load_state("networkidle")
        page.screenshot(path="pos_login.png")

        # 2. NAVEGAÇÃO NOS MENUS
        # wait_for garante que o menu carregou (necessário no headless)
        page.get_by_text("Relatórios").nth(1).wait_for(state="visible", timeout=60000)
        page.get_by_text("Relatórios").nth(1).click()
        page.get_by_role("link", name="LEC", exact=True).click()
        page.get_by_text("Efetividade de Leitura Faturamento", exact=True).click()

        # 3. DATAS DINÂMICAS
        print("Preenchendo datas automaticamente...")
        data_hoje = datetime.now().strftime("%d/%m/%Y")

        page.locator("input[name=\"j_idt93:0:j_idt99\"]").click()
        page.keyboard.type(data_hoje)
        page.keyboard.press("Tab")
        page.keyboard.type(data_hoje)
        page.keyboard.press("Tab")

        # 4. FILTROS

        # --- EMPRESAS ---
        print("Selecionando Empresas...")
        for emp in ["PAULISTA", "PIRATININGA"]:
            page.keyboard.type(emp)
            page.keyboard.press("Enter")
        page.keyboard.press("Tab")

        # --- REGIONAIS ---
        print("Selecionando Regionais...")
        for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
            page.keyboard.type(reg)
            page.keyboard.press("Enter")
        page.keyboard.press("Tab")

        # --- CIDADES ---
        print("Selecionando Cidades...")
        cidades = [
            "BAURU", "BOTUCATU", "JAU", "MARILIA", "INDAIATUBA",
            "JUNDIAI [B]", "ITU", "MAIRINQUE", "SOROCABA [B]",
            "JUNDIAI [A]", "SALTO [A]", "SOROCABA [A]"
        ]
        for cid in cidades:
            page.keyboard.type(cid)
            page.keyboard.press("Enter")

        # 5. FINALIZAÇÃO
        print("Clicando em Gerar Background...")
        page.get_by_role("button", name="Gerar Background").click()

        page.wait_for_timeout(2000)
        page.screenshot(path="pos_gerar.png")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Processo concluído com sucesso!")

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")
        page.screenshot(path="erro_automacao.png")
        raise

    finally:
        context.close()
        browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)