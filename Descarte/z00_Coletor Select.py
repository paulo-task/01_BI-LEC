import os
import time
import re
import sys
import zipfile
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv(dotenv_path=".pass")

# Recupera os dados centralizados
usuario = os.getenv("CPFL_USER")
senha = os.getenv("CPFL_PASS")

# --- CONFIGURAÇÃO DE PASTAS E REGRAS ---
MAPEAMENTO = {
    "01": ["Efetividade de Leitura Faturamento", r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario", "EFL", "data_hora_full"],
    "02": ["Produtividade Diária Leiturista - Analítico", r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico", "Produtividade Diária Leiturista - Analítico", "data_dia"],
    "03": ["Inst. Não Liberadas Faturamento", r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario", "Inst. Não Liberadas Faturamento", "substituir_puro"],
    "06": ["Lista Impedimentos Aplicados", r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos", "Lista Impedimentos Aplicados", "data_mes"],
    "07": ["Relatório de Efetividade de Entrega de Contas (Prev X Entr)", r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas", "Relatório de Efetividade de Entrega de Contas (Prev X Entr)", "data_mes"]
}

PASTA_04 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario"
PASTA_05 = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico"

def tratar_arquivo(caminho_zip, pasta_destino, nome_base, regra):
    time.sleep(3)
    try:
        if not zipfile.is_zipfile(caminho_zip):
            print(f"Erro: ZIP inválido: {caminho_zip}")
            return

        with zipfile.ZipFile(caminho_zip, 'r') as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()
            
            if regra == "data_hora_full":
                prefixo_hoje = f"{nome_base} - {agora.strftime('%Y_%m_%d')}"
                for f in os.listdir(pasta_destino):
                    if f.startswith(prefixo_hoje) and f.endswith(ext):
                        try: os.remove(os.path.join(pasta_destino, f))
                        except: pass
                nome_final = f"{prefixo_hoje}_{agora.strftime('%H-%M-%S')}{ext}"
            elif regra == "data_dia":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m_%d')}{ext}"
            elif regra == "data_mes":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m')}{ext}"
            else:
                nome_final = f"{nome_base}{ext}"

            caminho_novo = os.path.join(pasta_destino, nome_final)
            if os.path.exists(caminho_novo):
                try: os.remove(caminho_novo)
                except: pass

            z.extract(arquivo_interno, pasta_destino)
            caminho_extraido = os.path.join(pasta_destino, arquivo_interno)
            
            if os.path.abspath(caminho_extraido).lower() != os.path.abspath(caminho_novo).lower():
                os.rename(caminho_extraido, caminho_novo)
            
            print(f"--- Sucesso: '{nome_final}' atualizado.")
            
        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
            
    except Exception as e:
        print(f"Falha no tratamento: {e}")

def run(playwright: Playwright, escolha: str) -> None:
    user_dir = r"C:\Temp\dados_navegador_coletor"
    context = playwright.chromium.launch_persistent_context(user_dir, headless=False, no_viewport=True, slow_mo=300)
    page = context.pages[0]
    page.set_default_timeout(120000)

    try:
        print("\n[LOGIN] Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", wait_until="networkidle")
        
        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=10000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()
            page.wait_for_load_state("networkidle")

        print("[NAVEGAÇÃO] Acessando Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_selector("tbody tr", state="visible", timeout=120000)
        time.sleep(8)

        linhas = page.locator("tbody tr").all()

        # --- FILTRO DE EXECUÇÃO ---
        exec_geral = []
        exec_visitadas = False

        if escolha == "00":
            exec_geral = ["01", "02", "03", "06", "07"]
            exec_visitadas = True
        elif escolha in ["04", "05"]:
            exec_visitadas = True
        else:
            exec_geral = [escolha]

        # PASSO 1: MAPEAMENTO GERAL
        for id_rel in exec_geral:
            if id_rel in MAPEAMENTO:
                nome_filtro, pasta, n_base, regra = MAPEAMENTO[id_rel]
                
                maior_dt = None
                linha_alvo = None
                for linha in linhas:
                    t = linha.inner_text()
                    if nome_filtro.lower() in t.lower() and "Concluído" in t:
                        m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                        if m:
                            dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                            if maior_dt is None or dt > maior_dt:
                                maior_dt = dt
                                linha_alvo = linha
                
                if linha_alvo:
                    print(f"\n[DOWNLOAD] Baixando: {nome_filtro}")
                    os.makedirs(pasta, exist_ok=True)
                    with page.expect_download(timeout=180000) as dw:
                        linha_alvo.locator("button, a, img").filter(has_not_text=re.compile(r"Excluir|Remover|Exclui", re.I)).first.click()
                    tmp = os.path.join(pasta, f"temp_{id_rel}.zip")
                    dw.value.save_as(tmp)
                    tratar_arquivo(tmp, pasta, n_base, regra)

        # PASSO 2: NÃO VISITADAS (04 e 05)
        if exec_visitadas:
            todas_concluidas = []
            for linha in linhas:
                t = linha.inner_text()
                if "Instalações Não Visitadas" in t and "Concluído" in t:
                    m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    if m:
                        dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                        todas_concluidas.append({"dt": dt, "linha": linha})
            
            if len(todas_concluidas) >= 2:
                todas_concluidas.sort(key=lambda x: x['dt'], reverse=True)
                os_dois = todas_concluidas[:2]
                os_dois.sort(key=lambda x: x['dt'])
                
                # Se for escolha específica "04", ou ciclo completo
                if escolha in ["00", "04"]:
                    print(f"\n[DOWNLOAD] Baixando Diário (04): {os_dois[0]['dt']}")
                    os.makedirs(PASTA_04, exist_ok=True)
                    with page.expect_download() as dw:
                        os_dois[0]['linha'].locator("button, a, img").filter(has_not_text=re.compile(r"Excluir|Remover|Exclui", re.I)).first.click()
                    p04_zip = os.path.join(PASTA_04, "temp_04.zip")
                    dw.value.save_as(p04_zip)
                    tratar_arquivo(p04_zip, PASTA_04, "Instalações Não Visitadas", "data_dia")

                # Se for escolha específica "05", ou ciclo completo
                if escolha in ["00", "05"]:
                    print(f"\n[DOWNLOAD] Baixando Histórico (05): {os_dois[1]['dt']}")
                    os.makedirs(PASTA_05, exist_ok=True)
                    with page.expect_download() as dw:
                        os_dois[1]['linha'].locator("button, a, img").filter(has_not_text=re.compile(r"Excluir|Remover|Exclui", re.I)).first.click()
                    p05_zip = os.path.join(PASTA_05, "temp_05.zip")
                    dw.value.save_as(p05_zip)
                    tratar_arquivo(p05_zip, PASTA_05, "Instalações Não Visitadas", "data_dia")

    except Exception as e:
        print(f"Erro: {e}")
    finally:
        context.close()

if __name__ == "__main__":
    print("="*40)
    print("      COLETOR DE RELATÓRIOS CPFL")
    print("="*40)
    print("01 - ELF (Efetividade Leitura)")
    print("02 - PDL Analítico")
    print("03 - Inst. Não Liberadas Faturamento")
    print("04 - Não Visitadas (Diário)")
    print("05 - Não Visitadas (Histórico)")
    print("06 - Lista de Impedimentos")
    print("07 - Efetividade Entrega (Prev x Entr)")
    print("00 - CICLO COMPLETO")
    print("-"*40)
    
    op = input("Escolha uma opção: ").strip()
    
    if op in ["01", "02", "03", "04", "05", "06", "07", "00"]:
        with sync_playwright() as playwright:
            run(playwright, op)
    else:
        print("Opção inválida. Fechando...")