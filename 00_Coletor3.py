import os
import time
import re
import sys
import zipfile
import platform
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# =================================================================
# 1. TRATAMENTO DE VARIÁVEIS E CREDENCIAIS
# =================================================================
# Força a busca do arquivo .pass na mesma pasta deste script
diretorio_script = os.path.dirname(os.path.abspath(__file__))
caminho_pass = os.path.join(diretorio_script, ".pass")

if os.path.exists(caminho_pass):
    load_dotenv(dotenv_path=caminho_pass)
    print(f"✔ Arquivo .pass detectado em: {caminho_pass}")
else:
    load_dotenv() # Tenta carregar do ambiente (GitHub Actions)
    print("⚠️ Arquivo .pass não encontrado localmente. Usando variáveis de ambiente do sistema.")

# Fallback para string vazia "" evita o erro 'missing 1 required positional argument: value'
usuario = os.getenv("CPFL_USER", "")
senha = os.getenv("CPFL_PASS", "")

if not usuario or not senha:
    print("❌ ERRO CRÍTICO: Usuário ou Senha não carregados. Verifique o arquivo .pass.")
    # sys.exit() # Descomente para parar o script se as senhas falharem

# =================================================================
# 2. MAPEAMENTO DE PASTAS DINÂMICO
# =================================================================
def get_path(win_path):
    if platform.system() == "Windows":
        return win_path
    # No Linux (GitHub), remove C:\ e usa /tmp/downloads
    folder_name = win_path.split("\\")[-1]
    return os.path.join("/tmp/downloads", folder_name)

MAPEAMENTO = {
    "Efetividade de Leitura Faturamento": [get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario"), "EFL", "data_dia_unico"],
    "Produtividade Diária Leiturista - Analítico": [get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico"), "Produtividade Diária Leiturista - Analítico", "data_dia"],
    "Inst. Não Liberadas Faturamento": [get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario"), "Inst. Não Liberadas Faturamento", "substituir_puro"],
    "Lista Impedimentos Aplicados": [get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos"), "Lista Impedimentos Aplicados", "data_mes"],
    "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": [get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas"), "Relatório de Efetividade de Entrega de Contas (Prev X Entr)", "data_mes"]
}

PASTA_04 = get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario")
PASTA_05 = get_path(r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico")

# =================================================================
# 3. FUNÇÕES DE SUPORTE
# =================================================================
def tratar_arquivo(caminho_zip, pasta_destino, nome_base, regra):
    """Extrai o ZIP e aplica as regras de nomenclatura."""
    time.sleep(3)
    try:
        os.makedirs(pasta_destino, exist_ok=True)
        if not zipfile.is_zipfile(caminho_zip):
            print(f"Erro: ZIP inválido: {caminho_zip}")
            return

        with zipfile.ZipFile(caminho_zip, 'r') as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            agora = datetime.now()
            
            if regra == "data_dia_unico":
                prefixo_dia = f"{nome_base} - {agora.strftime('%Y_%m_%d')}"
                for f in os.listdir(pasta_destino):
                    if f.startswith(prefixo_dia) and f.endswith(ext):
                        try: os.remove(os.path.join(pasta_destino, f))
                        except: pass
                nome_final = f"{prefixo_dia}{ext}"
            elif regra == "data_dia":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m_%d')}{ext}"
            elif regra == "data_mes":
                nome_final = f"{nome_base}_{agora.strftime('%Y_%m')}{ext}"
            else:
                nome_final = f"{nome_base}{ext}"

            caminho_novo = os.path.join(pasta_destino, nome_final)
            z.extract(arquivo_interno, pasta_destino)
            caminho_extraido = os.path.join(pasta_destino, arquivo_interno)
            
            if os.path.abspath(caminho_extraido).lower() != os.path.abspath(caminho_novo).lower():
                if os.path.exists(caminho_novo): os.remove(caminho_novo)
                os.rename(caminho_extraido, caminho_novo)
            
            print(f"✔ Arquivo tratado: '{nome_final}'")
            
        if os.path.exists(caminho_zip): os.remove(caminho_zip)
    except Exception as e:
        print(f"Erro no tratamento do arquivo: {e}")

# =================================================================
# 4. CORE DO COLETOR (PLAYWRIGHT)
# =================================================================
def run(playwright: Playwright, busca=None) -> None:
    if platform.system() == "Windows":
        user_dir = r"C:\Temp\dados_navegador_coletor"
        is_headless = False
    else:
        user_dir = "/tmp/playwright_coletor"
        is_headless = True

    agora_str = datetime.now().strftime('%H:%M:%S')
    print(f"[{agora_str}] Iniciando Coletor (Headless: {is_headless})")

    context = playwright.chromium.launch_persistent_context(
        user_dir, 
        headless=is_headless, 
        no_viewport=True, 
        slow_mo=0 if is_headless else 500
    )
    page = context.pages[0]
    page.set_default_timeout(300000) 

    try:
        print("Acessando Portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login")
        
        # Login automático com validação
        user_field = page.get_by_role("textbox", name="Usuário")
        if user_field.is_visible(timeout=15000):
            user_field.fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()

        print("Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        
        page.wait_for_selector("tbody tr", state="visible", timeout=180000)
        time.sleep(10 if is_headless else 30) 

        linhas = page.locator("tbody tr").all()

        # --- PARTE 1: RELATÓRIOS GERAIS ---
        for nome_rel, conf in MAPEAMENTO.items():
            if busca and not any(t.lower() in nome_rel.lower() for t in busca):
                continue
            
            maior_dt = None
            linha_alvo = None
            for linha in linhas:
                t = linha.inner_text()
                if nome_rel.lower() in t.lower() and "Concluído" in t:
                    m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    if m:
                        dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
                        if maior_dt is None or dt > maior_dt:
                            maior_dt = dt
                            linha_alvo = linha
            
            if linha_alvo:
                print(f"Baixando: {nome_rel}")
                os.makedirs(conf[0], exist_ok=True)
                with page.expect_download(timeout=300000) as dw_info:
                    # Filtra para não clicar no botão excluir acidentalmente
                    linha_alvo.locator("button, a, img").filter(has_not_text=re.compile(r"Excluir|Remover", re.I)).first.click()
                
                tmp_path = os.path.join(conf[0], f"temp_{conf[1][:5]}.zip")
                dw_info.value.save_as(tmp_path)
                tratar_arquivo(tmp_path, conf[0], conf[1], conf[2])

        # --- PARTE 2: NÃO VISITADAS (04 e 05) ---
        if not busca or any("visitada" in b.lower() for b in busca):
            cand_04 = []; cand_05 = []
            for linha in linhas:
                t = linha.inner_text()
                if "Instalações Não Visitadas" in t and "Concluído" in t:
                    m_data = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", t)
                    m_params = re.findall(r"Dt(?:Ini|Fim)PrevLeitura=(\d{2}/\d{2}/\d{4})", t)
                    if m_data and len(m_params) >= 2:
                        dt_inclusao = datetime.strptime(m_data.group(1), "%d/%m/%Y %H:%M:%S")
                        if m_params[0] != m_params[1]: cand_04.append({"dt": dt_inclusao, "linha": linha})
                        else: cand_05.append({"dt": dt_inclusao, "linha": linha})

            if cand_04:
                cand_04.sort(key=lambda x: x['dt'], reverse=True)
                alvo = cand_04[0]
                print(f"Baixando Diário (04): {alvo['dt']}")
                with page.expect_download(timeout=300000) as dw:
                    alvo['linha'].locator("button, a, img").filter(has_not_text=re.compile(r"Excluir", re.I)).first.click()
                path_zip = os.path.join(PASTA_04, "temp_04.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_04, "Instalações Não Visitadas", "data_dia")

            if cand_05:
                cand_05.sort(key=lambda x: x['dt'], reverse=True)
                alvo = cand_05[0]
                print(f"Baixando Histórico (05): {alvo['dt']}")
                with page.expect_download(timeout=300000) as dw:
                    alvo['linha'].locator("button, a, img").filter(has_not_text=re.compile(r"Excluir", re.I)).first.click()
                path_zip = os.path.join(PASTA_05, "temp_05.zip")
                dw.value.save_as(path_zip)
                tratar_arquivo(path_zip, PASTA_05, "Instalações Não Visitadas", "data_dia")

    except Exception as e:
        print(f"❌ Erro Crítico: {e}")
        try:
            page.screenshot(path="erro_coletor.png")
            print("📸 Screenshot do erro salva como 'erro_coletor.png'")
        except: pass
    finally:
        context.close()

if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    with sync_playwright() as playwright:
        run(playwright, args)