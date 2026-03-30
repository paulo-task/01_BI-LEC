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

def get_primeiro_dia_util():
    # Pega o dia 1 do mês atual
    data = datetime.now().replace(day=1)
    # 5 = Sábado, 6 = Domingo. Pula para segunda-feira se necessário.
    if data.weekday() == 5:
        data = data + timedelta(days=2)
    elif data.weekday() == 6:
        data = data + timedelta(days=1)
    return data.strftime("%d/%m/%Y")

def run(playwright: Playwright) -> None:
    # slow_mo garante que o site acompanhe a digitação e cliques
    browser = playwright.chromium.launch(headless=False, slow_mo=600)
    context = browser.new_context()
    page = context.new_page()
    
    # 1. ACESSO E LOGIN
    page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
    page.get_by_role("textbox", name="Usuário").fill(usuario)
    page.get_by_role("textbox", name="Senha").fill(senha)
    page.get_by_role("button", name="Login").click()
    
    page.wait_for_load_state("networkidle")

    # 2. NAVEGAÇÃO
    page.get_by_text("Relatórios").nth(1).click()
    page.get_by_role("link", name="LEC", exact=True).click()
    page.get_by_text("Relatório de Efetividade de").click()

    # 3. DATAS DINÂMICAS
    data_inicio = get_primeiro_dia_util()
    data_fim = datetime.now().strftime("%d/%m/%Y")
    
    print(f"Relatório de Efetividade: {data_inicio} até {data_fim}")

    # Preenche Data Início (j_idt93:0)
    page.locator("input[name=\"j_idt93:0:j_idt99\"]").click()
    page.keyboard.type(data_inicio)
    page.keyboard.press("Tab")

    # Preenche Data Fim (j_idt93:1)
    page.keyboard.type(data_fim)
    page.keyboard.press("Tab")

    # 4. FILTROS (EMPRESA -> REGIONAL -> CIDADE)
    # Usando clique no texto exato para garantir a seleção correta
    
    # Empresas
    for emp in ["PAULISTA", "PIRATININGA"]:
        page.keyboard.type(emp)
        page.get_by_text(emp, exact=True).first.click()
    page.keyboard.press("Tab")
    
    # Regionais (Unidade de Negócio)
    for reg in ["PAULISTA-NOROESTE", "PIRATININGA-OESTE"]:
        page.keyboard.type(reg)
        page.get_by_text(reg, exact=True).first.click()
    page.keyboard.press("Tab")

    # Cidades
    cidades = [
        "BAURU", "BOTUCATU", "JAU", "MARILIA", "INDAIATUBA", 
        "JUNDIAI [B]", "ITU", "MAIRINQUE", "SOROCABA [B]",
        "JUNDIAI [A]", "SALTO [A]", "SOROCABA [A]"
    ]
    
    for cid in cidades:
        page.keyboard.type(cid)
        try:
            # Tenta o clique exato para não confundir Jundiaí [A] com [B]
            page.get_by_text(cid, exact=True).first.click()
        except:
            page.keyboard.press("Enter")
            
    page.keyboard.press("Tab")

    # 5. FINALIZAÇÃO
    print("Gerando Background...")
    page.get_by_role("button", name="Gerar Background").click()

    page.wait_for_timeout(3000)
    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)