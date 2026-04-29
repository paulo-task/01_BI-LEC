import os
import pandas as pd
import calendar
import requests
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from urllib.parse import urlparse
from playwright.sync_api import Playwright, sync_playwright

# --- DETECÇÃO DE AMBIENTE ---
IS_GITHUB = os.getenv('GITHUB_ACTIONS') == 'true'
FUSO_SP = ZoneInfo('America/Sao_Paulo')

if IS_GITHUB:
    usuario = os.getenv("CPFL_USER")
    senha = os.getenv("CPFL_PASS")
else:
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=".pass")
    usuario = os.getenv("CPFL_USER")
    senha = os.getenv("CPFL_PASS")

def get_diretorio_destino():
    if IS_GITHUB:
        path = os.path.join(os.getcwd(), 'downloads')
        os.makedirs(path, exist_ok=True)
        return path
    else:
        path = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\15_Planeja_Bases"
        os.makedirs(path, exist_ok=True)
        return path

def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo para o SharePoint via Microsoft Graph API"""
    try:
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID", "").strip()
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "").strip()
        SP_TENANT_ID = os.getenv("SP_TENANT_ID", "").strip()
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        if not SP_CLIENT_ID:
            print("   ⚠️ SharePoint: credenciais não configuradas")
            return False
        
        url_token = f'https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token'
        data = {
            'grant_type': 'client_credentials',
            'client_id': SP_CLIENT_ID,
            'client_secret': SP_CLIENT_SECRET,
            'scope': 'https://graph.microsoft.com/.default'
        }
        r = requests.post(url_token, data=data)
        r.raise_for_status()
        token = r.json()['access_token']
        
        parsed = urlparse(SITE_URL)
        host = parsed.netloc
        site_path = parsed.path.strip("/")
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url_site, headers=headers)
        r.raise_for_status()
        site_id = r.json()["id"]
        
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        r = requests.get(url_drives, headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        
        drive_id = None
        for drive in r.json().get('value', []):
            if drive.get('name') == 'Workspace':
                drive_id = drive.get('id')
                break
        
        if not drive_id:
            raise Exception("Workspace não encontrado")
        
        url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content"
        headers_upload = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        r = requests.put(url_upload, headers=headers_upload, data=conteudo_bytes)
        r.raise_for_status()
        
        print(f"   ☁️ Upload SharePoint: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"   ⚠️ Erro upload SharePoint: {e}")
        return False

def run(playwright: Playwright) -> None:
    # --- CONFIGURAÇÕES DE CAMINHO ---
    diretorio_destino = get_diretorio_destino()
    
    # --- AJUSTE DE DATAS ---
    hoje = datetime.now(FUSO_SP)
    data_atual_str = hoje.strftime("%d/%m/%Y") # Identico ao primeiro script
    
    # Data 1: Dia 16 do mês corrente
    data_inicio_filtro = hoje.replace(day=16).strftime("%d/%m/%Y")
    
    # Data 2: Cálculo do último dia útil do mês corrente
    ultimo_dia_mes = calendar.monthrange(hoje.year, hoje.month)[1]
    dt_fim = hoje.replace(day=ultimo_dia_mes)
    
    # Se o último dia for Sábado(5) ou Domingo(6), volta para Sexta
    while dt_fim.weekday() > 4:
        dt_fim -= timedelta(days=1)
    
    data_fim_filtro = dt_fim.strftime("%d/%m/%Y")
    
    # --- AJUSTE DE NOME: Deve ser identico ao primeiro script para mesclar ---
    data_nome_arq = hoje.strftime("%Y_%m_%d")
    nome_csv = f"Planejamento_Base_{data_nome_arq}.csv"
    caminho_final = os.path.join(diretorio_destino, nome_csv)

    # --- VERIFICAÇÃO DE PERMISSÃO ---
    if os.path.exists(caminho_final):
        try:
            with open(caminho_final, 'a') as f: pass 
        except PermissionError:
            print(f"\n[ERRO CRÍTICO] O arquivo {nome_csv} está aberto! Feche o Excel.")
            return

    browser = playwright.chromium.launch(headless=IS_GITHUB)
    context = browser.new_context()
    page = context.new_page()
    
    # --- CABEÇALHO IDENTICO AO PRIMEIRO SCRIPT ---
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

        print("Abrindo menus...")
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
                print(f"Preenchendo período Q2: {data_inicio_filtro} a {data_fim_filtro}")
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
                        info.append(data_atual_str) # Adiciona DT_RELATORIO identico ao script 1
                        dados_totais.append(info)

            page.locator(".rc-tree-checkbox.rc-tree-checkbox-checked").click(force=True)
            page.wait_for_timeout(500)

        # --- LÓGICA DE MESCLAGEM (APPEND SEM CABEÇALHO) ---
        if dados_totais:
            df_novo = pd.DataFrame(dados_totais, columns=CABECALHOS)
            
            if not os.path.exists(caminho_final):
                df_novo.to_csv(caminho_final, sep=';', index=False, encoding='latin1')
                print(f"Novo arquivo criado: {nome_csv}")
            else:
                df_novo.to_csv(caminho_final, sep=';', index=False, encoding='latin1', mode='a', header=False)
                print(f"Dados do período Q2 mesclados abaixo dos dados do Q1 em {nome_csv}.")
            
            # Upload para SharePoint (apenas no GitHub Actions)
            if IS_GITHUB and os.path.exists(caminho_final):
                with open(caminho_final, 'rb') as f:
                    conteudo = f.read()
                upload_to_sharepoint(conteudo, nome_csv, "BI_LEC/15_Planeja_Bases")
        else:
            print("Nenhum dado extraído.")

    finally:
        context.close()
        browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright: run(playwright)