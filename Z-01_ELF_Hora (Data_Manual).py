import os
import re
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright, expect

def run(playwright: Playwright) -> None:
    # 1. PERGUNTA AS DATAS ANTES DE ABRIR O NAVEGADOR
    # Assim você não precisa correr para digitar enquanto o navegador carrega
    print("\n" + "="*30)
    print(" CONFIGURAÇÃO DE DATAS")
    print("="*30)
    
    # Se você apenas der Enter, ele assume a data de hoje como padrão
    data_padrao = datetime.now().strftime("%d/%m/%Y")
    
    data_inicio = input(f"Digite a DATA INÍCIO (DD/MM/AAAA) [Padrão {data_padrao}]: ") or data_padrao
    data_fim = input(f"Digite a DATA FIM (DD/MM/AAAA) [Padrão {data_padrao}]: ") or data_padrao

    # Inicia o navegador
    browser = playwright.chromium.launch(headless=False, slow_mo=500)
    context = browser.new_context()
    page = context.new_page()
    
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Iniciando automação para o período: {data_inicio} até {data_fim}")

    try:
        # 2. ACESSO E LOGIN
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        page.get_by_role("textbox", name="Usuário").fill("CT34979")
        page.get_by_role("textbox", name="Senha").fill("FABI@1992")
        page.get_by_role("button", name="Login").click()
        
        page.wait_for_load_state("networkidle")

        # 3. NAVEGAÇÃO NOS MENUS
        page.get_by_text("Relatórios").nth(1).click()
        page.get_by_role("link", name="LEC", exact=True).click()
        page.get_by_text("Efetividade de Leitura Faturamento", exact=True).click()
        
        # 4. PREENCHIMENTO DAS DATAS ESCOLHIDAS
        print(f"Preenchendo datas: {data_inicio} e {data_fim}...")
        
        # Clica no primeiro campo de data e digita a data início
        page.locator("input[name=\"j_idt93:0:j_idt99\"]").click()
        page.keyboard.type(data_inicio)
        page.keyboard.press("Tab")
        
        # Digita no segundo campo a data fim
        page.keyboard.type(data_fim)
        page.keyboard.press("Tab")
        
        # 5. PREENCHIMENTO DOS FILTROS
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
            
        # 6. FINALIZAÇÃO
        print("Clicando em Gerar Background...")
        page.get_by_role("button", name="Gerar Background").click()
        
        page.wait_for_timeout(2000)
        print("Processo concluído com sucesso!")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    context.close()
    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)