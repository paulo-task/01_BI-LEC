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

def is_github_actions():
    """Detecta se está rodando no GitHub Actions"""
    return os.getenv("GITHUB_ACTIONS") == "true"

def get_headless():
    """Retorna True para headless no GitHub, False para Windows local"""
    if is_github_actions():
        return True
    return platform.system() != "Windows"

def run(playwright: Playwright) -> None:
    headless_mode = get_headless()
    ambiente = "GitHub Actions" if is_github_actions() else "Windows Local"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Ambiente: {ambiente} | Modo headless: {headless_mode}")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Instalações Não Visitadas - Histórico...")

    if not usuario or not senha:
        print("❌ ERRO: Usuário ou Senha não encontrados nas variáveis de ambiente!")
        return

    browser = playwright.chromium.launch(headless=headless_mode, slow_mo=500)
    context = browser.new_context()
    page = context.new_page()

    try:
        # 1. ACESSO E LOGIN
        print("🌐 Fazendo login...")
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
        print("📊 Navegando para Instalações Não Visitadas...")
        menu_relatorios = page.get_by_text("Relatórios").nth(1)
        menu_relatorios.wait_for(state="visible", timeout=60000)
        menu_relatorios.click()
        
        # Atraso estratégico para aguardar a abertura visual do dropdown do menu
        page.wait_for_timeout(2000)
        
        page.get_by_role("link", name="LEC", exact=True).click()
        page.get_by_text("Instalações Não Visitadas").click()

        # 3. FILTROS (EMPRESA -> REGIONAL -> CIDADE)
        print("🏢 Preenchendo Filtros de Localidade...")
        page.locator(".selectize-input").first.click()

        # Empresas
        print("   Selecionando Empresas...")
        for emp in ["PAULISTA", "PIRATININGA"]:
            page.keyboard.type(emp)
            page.get_by_text(emp, exact=True).first.click()
        page.keyboard.press("Tab")

        # Regionais
        print("   Selecionando Regionais...")
        for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
            page.keyboard.type(reg)
            page.get_by_text(reg, exact=True).first.click()
        page.keyboard.press("Tab")

        # Cidades
        print("   Selecionando Cidades...")
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

        # 4. GRUPO SERVIÇO (NOVO CAMPO)
        print("⚙️ Selecionando Grupo Serviço: BT-Baixa Tensão...")
        # Aguarda o campo ficar visível
        page.wait_for_selector("#rel-parametro-27", timeout=10000)
        # Seleciona a opção "BT"
        page.select_option("#rel-parametro-27", value="BT")
        print("   ✅ Grupo Serviço selecionado: BT-Baixa Tensão")

        # 5. DATAS (Data Atual em ambos os campos)
        data_atual = datetime.now().strftime("%d/%m/%Y")
        print(f"📅 Preenchendo Datas com o dia de hoje: {data_atual}")

        page.locator('input[name="j_idt93:8:j_idt99"]').click()
        page.keyboard.type(data_atual)
        page.keyboard.press("Tab")
        page.keyboard.type(data_atual)
        page.keyboard.press("Tab")

        # 6. GERAR RELATÓRIO
        print("🚀 Clicando em Gerar Background...")
        page.get_by_role("button", name="Gerar Background").click()
        page.wait_for_timeout(3000)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ Sucesso! O relatório está sendo gerado.")

    except Exception as e:
        print(f"❌ Erro durante a execução: {e}")
        raise

    finally:
        context.close()
        browser.close()


if __name__ == "__main__":
    print(f"⏰ Início: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with sync_playwright() as playwright:
        run(playwright)
    print(f"🏁 Fim: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")