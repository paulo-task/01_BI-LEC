import os
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

def run(playwright: Playwright) -> None:
    # --- CONFIGURAÇÕES DE CAMINHO ---
    diretorio_destino = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\15_Planeja_Bases"
    if not os.path.exists(diretorio_destino): os.makedirs(diretorio_destino)
    
    # --- AJUSTE DE DATAS ---
    hoje = datetime.now()
    data_atual_str = hoje.strftime("%d/%m/%Y") # Para a coluna DT_RELATORIO
    
    # Cálculo do primeiro dia útil do mês
    dt_inicio = hoje.replace(day=1)
    while dt_inicio.weekday() > 4: # Se for Sábado(5) ou Domingo(6), pula
        dt_inicio += timedelta(days=1)
    
    data_inicio_filtro = dt_inicio.strftime("%d/%m/%Y") # Data 1
    data_fim_filtro = hoje.replace(day=15).strftime("%d/%m/%Y") # Data 2
    
    # --- NOVO AJUSTE DE NOME: Planejamento_Base_AAAA_MM_DD.csv ---
    data_nome_arq = hoje.strftime("%Y_%m_%d")
    nome_csv = f"Planejamento_Base_{data_nome_arq}.csv"
    caminho_final = os.path.join(diretorio_destino, nome_csv)

    # --- REMOVIDA A LÓGICA DE EXCLUSÃO (os.remove) ---
    # Apenas verificamos se o arquivo está aberto caso queiramos sobrescrever o do mesmo dia
    if os.path.exists(caminho_final):
        try:
            # Testamos abertura para evitar PermissionError se o arquivo do dia estiver aberto
            with open(caminho_final, 'a') as f: pass 
        except PermissionError:
            print(f"\n[ERRO CRÍTICO] O arquivo {nome_csv} já existe e está aberto!")
            return

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    
    # --- ADICIONADO DT_RELATORIO AO CABEÇALHO ---
    CABECALHOS = ["TAREFA", "STATUS", "UNIDADE", "DESCRIÇÃO", "TIPO", "MINICIPIO", "DT PREVISTA", "DT LIMITE", "DT PLANEJA", "AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA", "T. TELEMED", "T. DISTRIB", "TP ATIVIDADE", "QTD FAT", "QTD FAT DISTRIB", "QTD FAT ENTREGUES", "QTD FAT DEVOLVIDAS", "NOME_BASE", "DT_RELATORIO"]
    NOMES_BASES = ["BAURU", "BOTUCATU", "JAU", "MARILIA", "INDAIATUBA", "JUNDIAI", "ITU", "MAIRINQUE", "SOROCABA"]
    
    SELETORES = [
        "li:nth-child(2) > .rc-tree-child-tree > li:nth-child(2) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(3) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(4) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(6) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(2) > ul > .rc-tree-treenode-switcher-open > .rc-tree-child-tree > li > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(2) > ul > .rc-tree-treenode-switcher-open > .rc-tree-child-tree > li:nth-child(4) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(2) > ul > .rc-tree-treenode-switcher-open > .rc-tree-child-tree > li:nth-child(2) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(2) > ul > .rc-tree-treenode-switcher-open > .rc-tree-child-tree > li:nth-child(5) > .rc-tree-node > .rc-tree-checkbox", 
        "li:nth-child(2) > ul > .rc-tree-treenode-switcher-open > .rc-tree-child-tree > li:nth-child(8) > .rc-tree-node > .rc-tree-checkbox"
    ]

    dados_totais = []

    try:
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        page.get_by_role("textbox", name="Usuário").fill(usuario)
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()
        page.get_by_text("Planejamento / Operação").nth(1).click()

        print("Abrindo menus da árvore...")
        page.locator(".rc-tree-switcher").first.click()
        page.locator("li:nth-child(2) > .rc-tree-node > .rc-tree-switcher").first.click()
        page.locator(".rc-tree > li:nth-child(2) > .rc-tree-node > .rc-tree-switcher").click() 
        page.locator(".rc-tree > li:nth-child(2) > .rc-tree-child-tree > li:nth-child(2) > .rc-tree-node > .rc-tree-switcher").click()
        
        page.locator(".switch").first.click()
        page.locator("span:nth-child(7) > .switch").click()

        for i in range(len(SELETORES)):
            nome_da_base = NOMES_BASES[i]
            print(f"Processando: {nome_da_base}...")

            page.locator(SELETORES[i]).first.click()

            if i == 0:
                print(f"Preenchendo período: {data_inicio_filtro} a {data_fim_filtro}")
                for n in [1, 2]:
                    page.get_by_role("textbox", name="Data dd/mm/yyyy").nth(n).click()
                    page.keyboard.press("Control+A")
                    page.keyboard.press("Backspace")
                    valor_data = data_inicio_filtro if n == 1 else data_fim_filtro
                    page.get_by_role("textbox", name="Data dd/mm/yyyy").nth(n).fill(valor_data)
                page.keyboard.press("Enter")

            page.wait_for_timeout(3000)
            page.get_by_title("Filtrar por data prevista").click()
            page.wait_for_timeout(12000) 

            linhas = page.locator("tbody tr").all()
            for linha in linhas:
                colunas = linha.locator("td").all_text_contents()
                if len(colunas) >= 21:
                    info = [txt.strip() for txt in colunas[2:21]]
                    if any(info[:3]):
                        info.append(nome_da_base)
                        # ADICIONADO: Preenchimento da data atual na nova coluna
                        info.append(data_atual_str) 
                        dados_totais.append(info)

            page.locator(".rc-tree-checkbox.rc-tree-checkbox-checked").click()
            page.wait_for_timeout(500)

        if dados_totais:
            pd.DataFrame(dados_totais, columns=CABECALHOS).to_csv(caminho_final, sep=';', index=False, encoding='latin1')
            print(f"\n[SUCESSO] Relatório {nome_csv} gerado.")
        else:
            print("Nenhum dado extraído.")

    finally:
        context.close()
        browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright: run(playwright)