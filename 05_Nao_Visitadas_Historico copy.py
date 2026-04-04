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
        f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Instalações Não Visitadas - Histórico..."
    )

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

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

        # 2. NAVEGAÇÃO
        page.get_by_text("Relatórios").nth(1).wait_for(state="visible", timeout=60000)
        page.get_by_text("Relatórios").nth(1).click()
        page.get_by_role("link", name="LEC", exact=True).click()
        page.get_by_text("Instalações Não Visitadas").click()

        # 3. FILTROS (EMPRESA -> REGIONAL -> CIDADE)
        print("Preenchendo Filtros de Localidade...")
        page.locator(".selectize-input").first.click()

        # Empresas
        for emp in ["PAULISTA", "PIRATININGA"]:
            page.keyboard.type(emp)
            page.get_by_text(emp, exact=True).first.click()
        page.keyboard.press("Tab")

        # Regionais
        for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
            page.keyboard.type(reg)
            page.get_by_text(reg, exact=True).first.click()
        page.keyboard.press("Tab")

        # Cidades
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

        # 4. DATAS (Data Atual em ambos os campos)
        data_atual = datetime.now().strftime("%d/%m/%Y")
        print(f"Preenchendo Datas com o dia de hoje: {data_atual}")

        page.locator('input[name="j_idt93:8:j_idt99"]').click()
        page.keyboard.type(data_atual)
        page.keyboard.press("Tab")
        page.keyboard.type(data_atual)
        page.keyboard.press("Tab")

        # 5. GERAR RELATÓRIO
        print("Clicando em Gerar Background...")
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
