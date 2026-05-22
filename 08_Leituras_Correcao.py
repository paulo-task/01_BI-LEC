import os
import re
from datetime import datetime, timedelta
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

load_dotenv(dotenv_path=".pass")

usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")


def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"


def get_headless():
    if is_github_actions():
        return True
    return False


def get_primeiro_dia_util():
    data = datetime.now().replace(day=1)
    if data.weekday() == 5:
        data = data + timedelta(days=2)
    elif data.weekday() == 6:
        data = data + timedelta(days=1)
    return data.strftime("%d/%m/%Y")


def run(playwright: Playwright) -> None:
    headless_mode = get_headless()
    ambiente = "GitHub Actions" if is_github_actions() else "Windows Local"
    print(f"Ambiente: {ambiente} | Modo headless: {headless_mode}")

    browser = playwright.chromium.launch(headless=headless_mode, slow_mo=600)
    context = browser.new_context(viewport={"width": 1920, "height": 1080})
    page = context.new_page()

    print("Fazendo login...")
    page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)

    usuario_input = page.get_by_role("textbox", name="Usuário")
    usuario_input.wait_for(state="visible", timeout=60000)
    usuario_input.fill(usuario)

    page.get_by_role("textbox", name="Senha").fill(senha)
    page.get_by_role("button", name="Login").click()
    page.wait_for_load_state("networkidle", timeout=60000)

    print("Navegando para Leituras de Correção...")
    menu_relatorios = page.get_by_text("Relatórios").nth(1)
    menu_relatorios.wait_for(state="visible", timeout=60000)
    menu_relatorios.click()
    page.wait_for_timeout(2000)

    page.get_by_text("LEC", exact=True).first.click()
    page.get_by_text("Leituras de Correção").click()
    page.wait_for_load_state("networkidle", timeout=60000)

    data_inicio = get_primeiro_dia_util()
    data_fim = datetime.now().strftime("%d/%m/%Y")
    print(f"Datas: {data_inicio} até {data_fim}")

    page.locator('input[name="j_idt93:0:j_idt99"]').wait_for(state="visible", timeout=60000)
    page.locator('input[name="j_idt93:0:j_idt99"]').click()
    page.keyboard.type(data_inicio)
    page.keyboard.press("Tab")
    page.keyboard.type(data_fim)
    page.keyboard.press("Tab")

    print("Selecionando Empresas...")
    page.locator(".selectize-input").first.click()
    page.get_by_text("PAULISTA", exact=True).click()
    page.get_by_text("PIRATININGA", exact=True).click()
    page.locator("#rel-parametro-25-selectized").press("Tab")

    print("Selecionando Regionais...")
    page.get_by_text("PAULISTA-NOROESTE", exact=True).click()
    page.get_by_text("PIRATININGA-OESTE", exact=True).click()
    page.locator("#rel-parametro-34-selectized").press("Tab")

    print("Selecionando Bases...")
    bases = [
        "BAURU [B]", "BOTUCATU [B]", "JAU [B]", "MARILIA [B]",
        "INDAIATUBA [B]", "JUNDIAI [B]", "ITU [B]", "MAIRINQUE [B]", "SOROCABA [B]",
    ]
    for base in bases:
        page.get_by_text(base, exact=True).click()
    page.locator("#rel-parametro-36-selectized").press("Tab")

    print("Configurando tensão...")
    page.locator("#rel-parametro-27").wait_for(state="visible", timeout=60000)
    page.locator("#rel-parametro-27").select_option("BT")

    print("Gerando relatório em background...")
    btn_gerar = page.get_by_role("button", name="Gerar Background")
    btn_gerar.wait_for(state="visible", timeout=60000)
    btn_gerar.scroll_into_view_if_needed()
    btn_gerar.click()

    print("Relatório solicitado com sucesso!")
    page.wait_for_timeout(3000)

    try:
        print("Deslogando do sistema para liberar a sessão...")
        page.get_by_role("menuitem", name=re.compile("Logout", re.IGNORECASE)).click(
            force=True, timeout=5000
        )
        page.wait_for_load_state("networkidle", timeout=5000)
        print("Sessão encerrada com sucesso.")
    except Exception as e:
        print(f"Aviso: Não foi possível deslogar automaticamente. {e}")

    context.close()
    browser.close()


if __name__ == "__main__":
    print(f"Início: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with sync_playwright() as playwright:
        run(playwright)
    print(f"Fim: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
