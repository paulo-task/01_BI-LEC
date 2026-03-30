import os
import time
import re
import zipfile
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# --- 1. CARREGAMENTO DE CONFIGURAÇÕES ---
# Centraliza a senha no arquivo .pass para facilitar mudanças futuras
load_dotenv(dotenv_path=".pass")

usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# --- 2. MAPEAMENTO DE PASTAS E RELATÓRIOS ---
CONFIG = {
    "PDL_Analitico": {
        "pasta": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
        "busca": "Produtividade Diária Leiturista - Analítico",
        "nome_final_base": "Produtividade Diária Leiturista - Analítico"
    },
    "Nao_Visitadas": {
        "pasta": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario",
        "busca": "Instalações Não Visitadas",
        "nome_final_base": "Instalações Não Visitadas"
    }
}

def tratar_arquivo(caminho_zip, pasta_destino, nome_base):
    """Extrai o ZIP, renomeia com a data atual e limpa o temporário."""
    try:
        if not zipfile.is_zipfile(caminho_zip):
            print(f"Erro: O arquivo {caminho_zip} não é um ZIP válido.")
            return
            
        with zipfile.ZipFile(caminho_zip, 'r') as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            z.extract(arquivo_interno, pasta_destino)
            
            nome_final = f"{nome_base}_{datetime.now().strftime('%Y_%m_%d')}{ext}"
            caminho_antigo = os.path.join(pasta_destino, arquivo_interno)
            caminho_novo = os.path.join(pasta_destino, nome_final)

            if os.path.exists(caminho_novo):
                os.remove(caminho_novo)
            os.rename(caminho_antigo, caminho_novo)
            print(f"✔ Sucesso: '{nome_final}' atualizado.")
            
        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
    except Exception as e:
        print(f"Erro no tratamento do arquivo: {e}")

def run(playwright: Playwright) -> None:
    # --- 3. CONFIGURAÇÃO DO NAVEGADOR ---
    user_dir = r"C:\Temp\dados_navegador_coletor"
    context = playwright.chromium.launch_persistent_context(
        user_dir, 
        headless=False, 
        no_viewport=True, 
        slow_mo=500 # Ações mais lentas para evitar bloqueios do portal
    )
    page = context.pages[0]

    # AJUSTE DE TEMPO 1: Timeout global de 3 minutos para qualquer comando
    page.set_default_timeout(180000) 

    try:
        if not usuario or not senha:
            print("Erro: Usuário ou Senha não encontrados no arquivo .pass")
            return

        print(f"[{datetime.now().strftime('%H:%M:%S')}] Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=90000)
        
        # Login automático
        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=15000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()
            page.wait_for_load_state("networkidle")

        print("Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        
        # AJUSTE DE TEMPO 2: Espera a tabela aparecer e dá tempo para os dados carregarem
        print("Aguardando carregamento da tabela de downloads...")
        page.wait_for_selector("tbody tr", state="visible", timeout=120000)
        time.sleep(25) # Tempo essencial para o portal CPFL renderizar todas as linhas
        
        linhas = page.locator("tbody tr").all()
        print(f"Total de linhas encontradas na tabela: {len(linhas)}")

        # --- 4. PROCESSAMENTO DOS RELATÓRIOS ---
        for chave, cfg in CONFIG.items():
            print(f"Buscando versão mais recente de: {cfg['busca']}...")
            
            maior_data_hora = None
            linha_alvo = None

            for linha in linhas:
                texto = linha.inner_text()
                
                # Filtra por nome e status concluído
                if cfg['busca'].lower() in texto.lower() and "Concluído" in texto:
                    match_data = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
                    
                    if match_data:
                        data_hora_str = match_data.group(1)
                        data_hora_dt = datetime.strptime(data_hora_str, "%d/%m/%Y %H:%M:%S")
                        
                        # Lógica para pegar sempre o download mais novo da lista
                        if maior_data_hora is None or data_hora_dt > maior_data_hora:
                            maior_data_hora = data_hora_dt
                            linha_alvo = linha

            if linha_alvo:
                print(f"-> Localizado: {maior_data_hora}. Iniciando download...")
                os.makedirs(cfg['pasta'], exist_ok=True)
                
                try:
                    # AJUSTE DE TEMPO 3: Timeout de download aumentado para 5 minutos
                    with page.expect_download(timeout=300000) as download_info:
                        # Clica no botão de download (evitando o botão excluir)
                        linha_alvo.locator("button, a, img").filter(
                            has_not_text=re.compile(r"Excluir|Remover|Exclui", re.I)
                        ).first.click()
                    
                    download = download_info.value
                    tmp_zip = os.path.join(cfg['pasta'], f"temp_{chave}.zip")
                    download.save_as(tmp_zip)
                    
                    tratar_arquivo(tmp_zip, cfg['pasta'], cfg['nome_final_base'])
                except Exception as e:
                    print(f"Atenção: Falha no download ou timeout para {chave}: {e}")
            else:
                print(f"Aviso: Não encontramos nenhum relatório '{cfg['busca']}' com status Concluído.")

    except Exception as e:
        print(f"Erro Crítico durante a execução: {e}")
    finally:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Finalizando sessão do navegador...")
        time.sleep(2)
        context.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)