import os
import re
import platform
from datetime import datetime, timedelta
from playwright.sync_api import Playwright, sync_playwright, expect
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

def is_github_actions():
    """Detecta se está rodando no GitHub Actions"""
    return os.getenv("GITHUB_ACTIONS") == "true"

def get_headless():
    """Retorna True para headless no GitHub, False para Windows local"""
    if is_github_actions():
        return True  # GitHub Actions: sem interface gráfica
    return platform.system() != "Windows"  # Windows: mostra janela

def get_primeiro_dia_util():
    # Pega o dia 1 do mês atual
    data = datetime.now().replace(day=1)
    # 5 = Sábado, 6 = Domingo. Se cair neles, pula para Segunda.
    if data.weekday() == 5:
        data = data + timedelta(days=2)
    elif data.weekday() == 6:
        data = data + timedelta(days=1)
    return data.strftime("%d/%m/%Y")

def run(playwright: Playwright) -> None:
    # Detecta ambiente e configura headless automaticamente
    headless_mode = get_headless()
    ambiente = "GitHub Actions" if is_github_actions() else "Windows Local"
    print(f"🏗️ Ambiente: {ambiente} | Modo headless: {headless_mode}")
    
    # slow_mo ajuda o site a processar as seleções sem pressa
    browser = playwright.chromium.launch(headless=headless_mode, slow_mo=600)
    context = browser.new_context()
    page = context.new_page()
    
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
    print("📊 Navegando para Lista Impedimentos Aplicados...")
    menu_relatorios = page.get_by_text("Relatórios").nth(1)
    menu_relatorios.wait_for(state="visible", timeout=60000)
    menu_relatorios.click()
    
    # Atraso estratégico para aguardar a abertura visual do dropdown do menu
    page.wait_for_timeout(2000)
    
    page.get_by_role("link", name="LEC", exact=True).click()
    page.get_by_text("Lista Impedimentos Aplicados", exact=True).click()

    # 3. DATAS (1º Útil e Atual)
    data_inicio = get_primeiro_dia_util()
    data_fim = datetime.now().strftime("%d/%m/%Y")
    
    print(f"📅 Datas: {data_inicio} até {data_fim}")

    # Preenche Data Início
    page.locator("input[name=\"j_idt93:0:j_idt99\"]").click()
    page.keyboard.type(data_inicio)
    page.keyboard.press("Tab")

    # Preenche Data Fim
    page.keyboard.type(data_fim)
    page.keyboard.press("Tab")

    # 4. FILTROS (EMPRESA -> BASE -> CIDADES)
    # Usando a técnica de digitar e clicar no texto exato para não errar a cidade
    
    # Empresas
    print("🏢 Selecionando Empresas...")
    for emp in ["PAULISTA", "PIRATININGA"]:
        page.keyboard.type(emp)
        page.get_by_text(emp, exact=True).first.click()
    page.keyboard.press("Tab")
    
    # Base (Regionais)
    print("📍 Selecionando Regionais...")
    for base in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
        page.keyboard.type(base)
        page.get_by_text(base, exact=True).first.click()
    page.keyboard.press("Tab")

    # Cidades
    print("🏙️ Selecionando Cidades...")
    cidades = [
        "BAURU", "BOTUCATU", "JAU", "MARILIA", "INDAIATUBA", 
        "JUNDIAI [B]", "ITU", "MAIRINQUE", "SOROCABA [B]"
    ]
    for cid in cidades:
        page.keyboard.type(cid)
        # Tenta clicar no texto exato. Se falhar, usa o Enter como reserva.
        try:
            page.get_by_text(cid, exact=True).first.click()
        except:
            page.keyboard.press("Enter")
    page.keyboard.press("Tab")

    # 5. TENSÃO E FINALIZAÇÃO
    print("⚙️ Configurando tensão...")
    page.locator("#rel-parametro-27").select_option("BT")
    
    print("🚀 Solicitando relatório em background...")
    page.get_by_role("button", name="Gerar Background").click()

    print("✅ Relatório solicitado com sucesso!")
    page.wait_for_timeout(3000)
    context.close()
    browser.close()

# Ponto de entrada
if __name__ == "__main__":
    print(f"⏰ Início: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with sync_playwright() as playwright:
        run(playwright)
    print(f"🏁 Fim: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")