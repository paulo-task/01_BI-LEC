import os
import re
from datetime import datetime, timedelta
from playwright.sync_api import Playwright, sync_playwright, expect
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

def run(playwright: Playwright) -> None:
    # --- CÁLCULO DAS DATAS ---
    hoje = datetime.now()
    dt_inicio = hoje.replace(day=1)
    while dt_inicio.weekday() > 4:
        dt_inicio += timedelta(days=1)
    data_inicio_str = dt_inicio.strftime("%d/%m/%Y")
    
    dt_amanha = hoje + timedelta(days=1)
    data_amanha_str = dt_amanha.strftime("%d/%m/%Y")

    # --- INÍCIO DO NAVEGADOR ---
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    
    page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)
    
    # Espera obrigatória de até 60s para a página estar pronta
    usuario_input = page.get_by_role("textbox", name="Usuário")
    usuario_input.wait_for(state="visible", timeout=60000)
    usuario_input.fill(usuario)
    
    page.get_by_role("textbox", name="Senha").fill(senha)
    page.get_by_role("button", name="Login").click()
    
    # Aguarda a tela pós-login carregar completamente
    page.wait_for_load_state("networkidle", timeout=60000)
    
    menu_relatorios = page.get_by_text("Relatórios").nth(1)
    menu_relatorios.wait_for(state="visible", timeout=60000)
    menu_relatorios.click()
    
    # Atraso estratégico para aguardar a abertura visual do dropdown do menu
    page.wait_for_timeout(2000)
    
    page.get_by_role("link", name="LEC", exact=True).click()
    page.get_by_text("Quantidade de Leituras Previstas por UL").click()

    # --- PREENCHIMENTO DAS DATAS ---
    # Data 1
    page.locator("input[name=\"j_idt93:0:j_idt99\"]").click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.locator("input[name=\"j_idt93:0:j_idt99\"]").fill(data_inicio_str)
    page.keyboard.press("Enter")

    # Data 2
    page.locator("input[name=\"j_idt93:1:j_idt99\"]").click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.locator("input[name=\"j_idt93:1:j_idt99\"]").fill(data_amanha_str)
    page.keyboard.press("Enter")

    # --- SELEÇÃO DE FILTROS (CORRIGIDO COM EXACT=TRUE) ---
    page.locator(".selectize-input").first.click()
    page.get_by_text("PAULISTA", exact=True).click()
    page.get_by_text("PIRATININGA", exact=True).click()
    page.locator("#rel-parametro-25-selectized").press("Tab")
    
    # Aqui estava o erro: adicionei exact=True
    page.get_by_text("PAULISTA-NOROESTE", exact=True).click()
    page.get_by_text("PIRATININGA-OESTE", exact=True).click()
    page.locator("#rel-parametro-34-selectized").press("Tab")
    
    # Bases
    bases = ["BAURU [B]", "BOTUCATU [B]", "JAU [B]", "MARILIA [B]", 
             "INDAIATUBA [B]", "JUNDIAI [B]", "ITU [B]", "MAIRINQUE [B]", "SOROCABA [B]"]
    
    for base in bases:
        # Usando exact=True aqui também por segurança
        page.get_by_text(base, exact=True).click()
        
    page.locator("#rel-parametro-36-selectized").press("Tab")
    page.locator("#rel-parametro-27").select_option("BT")
    
    page.get_by_role("button", name="Gerar Background").click()

    page.wait_for_timeout(3000)

    # Tentativa de Logout para liberar a sessão no servidor da CPFL
    try:
        print("Deslogando do sistema para liberar a sessão...")
        import re
        page.get_by_role("menuitem", name=re.compile("Logout", re.IGNORECASE)).click(force=True, timeout=5000)
        page.wait_for_load_state("networkidle", timeout=5000)
        print("Sessão encerrada com sucesso.")
    except Exception as e:
        print(f"Aviso: Não foi possível deslogar automaticamente. {e}")

    context.close()
    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)