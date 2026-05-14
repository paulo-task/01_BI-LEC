import os
import time
import re
import sys
import io
import zipfile
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# =================================================================
# 1. CARREGAR CREDENCIAIS
# =================================================================
diretorio_script = os.path.dirname(os.path.abspath(__file__))
caminho_pass = os.path.join(diretorio_script, ".pass")

if os.path.exists(caminho_pass):
    load_dotenv(dotenv_path=caminho_pass)
    print(f"✅ Arquivo .pass carregado de: {caminho_pass}")
else:
    load_dotenv()
    print("⚠️ Arquivo .pass não encontrado. Usando variáveis de ambiente.")

usuario = os.getenv("CPFL_USER", "")
senha = os.getenv("CPFL_PASS", "")

if not usuario or not senha:
    print("❌ ERRO: Credenciais CPFL não encontradas!")
    sys.exit(1)

print(f"✅ Credenciais CPFL carregadas")

# =================================================================
# 2. DETECTAR AMBIENTE
# =================================================================
def is_github_actions():
    return os.getenv("GITHUB_ACTIONS") == "true"

if is_github_actions():
    USE_HEADLESS = True
    print("🏗️ Ambiente: GitHub Actions (modo headless)")
else:
    USE_HEADLESS = False
    print("🖥️ Ambiente: Local (modo visível)")

# =================================================================
# 3. DEFINIÇÃO DOS RELATÓRIOS COM REGRAS DE NOMENCLATURA
# =================================================================

# Regras de nomenclatura personalizadas
def get_nome_arquivo(conf, ext):
    """Gera o nome do arquivo conforme as regras específicas"""
    agora = datetime.now()
    ano_mes = agora.strftime('%Y_%m')
    ano_mes_dia = agora.strftime('%Y_%m_%d')
    
    # Regras específicas por tipo de relatório
    if conf.get("regra_nome") == "ano_mes":
        # Ex: Nome_2026_04.csv
        return f"{conf['nome_base']}_{ano_mes}{ext}"
    
    elif conf.get("regra_nome") == "sem_data":
        # Ex: Nome.csv (sobrescreve)
        return f"{conf['nome_base']}{ext}"
    
    elif conf.get("regra_nome") == "data_dia_unico":
        # Ex: Nome - 2026_04_04.csv
        return f"{conf['nome_base']} - {ano_mes_dia}{ext}"
    
    else:  # padrão data_dia
        # Ex: Nome_2026_04_04.csv
        return f"{conf['nome_base']}_{ano_mes_dia}{ext}"

# Modo SIMPLES (3 relatórios)
RELATORIOS_SIMPLES = {
    "Efetividade de Leitura Faturamento": {
        "pasta_sp": "BI_LEC/01_ELF_Diario",
        "pasta_local": "01_ELF_Diario",
        "nome_base": "EFL",
        "regra_nome": "data_dia_unico",
        "filtro_grupo_servico": None
    },
    "Produtividade Diária Leiturista - Analítico": {
        "pasta_sp": "BI_LEC/02_PDL_Analitico",
        "pasta_local": "02_PDL_Analitico",
        "nome_base": "Produtividade Diária Leiturista - Analítico",
        "regra_nome": "data_dia",
        "filtro_grupo_servico": None
    },
    "Instalações Não Visitadas": {
        "pasta_sp": "BI_LEC/04_N_Visitado_Diario",
        "pasta_local": "04_N_Visitado_Diario",
        "nome_base": "Instalações Não Visitadas",
        "regra_nome": "data_dia",
        "filtro_grupo_servico": "vazio"
    }
}

# Modo COMPLETO (7 relatórios)
RELATORIOS_COMPLETOS = {
    "Efetividade de Leitura Faturamento": {
        "pasta_sp": "BI_LEC/01_ELF_Diario",
        "pasta_local": "01_ELF_Diario",
        "nome_base": "EFL",
        "regra_nome": "data_dia_unico",
        "filtro_grupo_servico": None,
        "nome_exibicao": "Efetividade de Leitura Faturamento"
    },
    "Produtividade Diária Leiturista - Analítico": {
        "pasta_sp": "BI_LEC/02_PDL_Analitico",
        "pasta_local": "02_PDL_Analitico",
        "nome_base": "Produtividade Diária Leiturista - Analítico",
        "regra_nome": "data_dia",
        "filtro_grupo_servico": None,
        "nome_exibicao": "Produtividade Diária Leiturista - Analítico"
    },
    "Instalações Não Visitadas (Diário)": {
        "pasta_sp": "BI_LEC/04_N_Visitado_Diario",
        "pasta_local": "04_N_Visitado_Diario",
        "nome_base": "Instalações Não Visitadas",
        "regra_nome": "data_dia",
        "filtro_grupo_servico": "vazio",
        "nome_exibicao": "Instalações Não Visitadas"
    },
    "Instalações Não Visitadas (Histórico)": {
        "pasta_sp": "BI_LEC/05_N_Visitado_Historico",
        "pasta_local": "05_N_Visitado_Historico",
        "nome_base": "Instalações Não Visitadas",
        "regra_nome": "data_dia",
        "filtro_grupo_servico": "BT",
        "nome_exibicao": "Instalações Não Visitadas"
    },
    "Lista Impedimentos Aplicados": {
        "pasta_sp": "BI_LEC/06_Impedimentos",
        "pasta_local": "06_Impedimentos",
        "nome_base": "Lista Impedimentos Aplicados",
        "regra_nome": "ano_mes",  # <-- Nome_2026_04.csv
        "filtro_grupo_servico": None,
        "nome_exibicao": "Lista Impedimentos Aplicados"
    },
    "Inst. Não Liberadas Faturamento": {
        "pasta_sp": "BI_LEC/03_N_Lib_Fat_Diario",
        "pasta_local": "03_N_Lib_Fat_Diario",
        "nome_base": "Inst. Não Liberadas Faturamento",
        "regra_nome": "sem_data",  # <-- Nome.csv (sobrescreve)
        "filtro_grupo_servico": None,
        "nome_exibicao": "Inst. Não Liberadas Faturamento"
    },
    "Relatório de Efetividade de Entrega de Contas (Prev X Entr)": {
        "pasta_sp": "BI_LEC/07_Entregas",
        "pasta_local": "07_Entregas",
        "nome_base": "Relatório de Efetividade de Entrega de Contas (Prev X Entr)",
        "regra_nome": "ano_mes",  # <-- Nome_2026_04.csv
        "filtro_grupo_servico": None,
        "nome_exibicao": "Relatório de Efetividade de Entrega de Contas (Prev X Entr)"
    }
}

# Pastas locais
PASTAS_LOCAIS = {
    "01_ELF_Diario": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
    "02_PDL_Analitico": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
    "03_N_Lib_Fat_Diario": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario",
    "04_N_Visitado_Diario": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario",
    "05_N_Visitado_Historico": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico",
    "06_Impedimentos": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos",
    "07_Entregas": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas"
}

# =================================================================
# 4. FUNÇÃO DE UPLOAD PARA SHAREPOINT
# =================================================================
def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo para o SharePoint"""
    try:
        import requests
        from urllib.parse import urlparse
        
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID", "").strip()
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "").strip()
        SP_TENANT_ID = os.getenv("SP_TENANT_ID", "").strip()
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        if not SP_CLIENT_ID:
            print("   ⚠️ SharePoint: credenciais não configuradas")
            return False
        
        # Obter token
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
        
        # Obter Site ID
        parsed = urlparse(SITE_URL)
        host = parsed.netloc
        site_path = parsed.path.strip("/")
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url_site, headers=headers)
        r.raise_for_status()
        site_id = r.json()["id"]
        
        # Obter Drive Workspace
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
        
        # Upload
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

# =================================================================
# 5. FUNÇÃO PARA SALVAR LOCALMENTE
# =================================================================
def salvar_localmente(conteudo_bytes, nome_arquivo, pasta_destino):
    """Salva arquivo localmente"""
    try:
        os.makedirs(pasta_destino, exist_ok=True)
        caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
        
        with open(caminho_arquivo, 'wb') as f:
            f.write(conteudo_bytes)
        
        print(f"   💾 Salvo localmente: {caminho_arquivo}")
        return True
    except Exception as e:
        print(f"   ❌ Erro ao salvar localmente: {e}")
        return False

# =================================================================
# 6. PROCESSAR ZIP
# =================================================================
def processar_zip(conteudo_zip, conf, destino_pasta):
    """Processa o ZIP e salva/envia usando as regras de nomenclatura"""
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo_zip)) as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            
            # Gera o nome do arquivo conforme as regras específicas
            nome_final = get_nome_arquivo(conf, ext)
            
            with z.open(arquivo_interno) as f:
                conteudo = f.read()
            
            if is_github_actions():
                return upload_to_sharepoint(conteudo, nome_final, destino_pasta)
            else:
                return salvar_localmente(conteudo, nome_final, destino_pasta)
            
    except Exception as e:
        print(f"   ❌ Erro processar ZIP: {e}")
        return False

# =================================================================
# 7. FUNÇÃO PARA ENCONTRAR LINHA DO RELATÓRIO
# =================================================================
def encontrar_linha_relatorio(linhas, nome_relatorio, filtro_grupo_servico=None):
    """
    Encontra a linha mais recente de um relatório com status "Concluído"
    
    filtro_grupo_servico:
    - None: não aplica filtro
    - "vazio": pega linhas que NÃO contém "&GrupoServico=BT"
    - "BT": pega linhas que contém "&GrupoServico=BT"
    """
    maior_dt = None
    linha_alvo = None
    info_adicional = ""
    
    for linha in linhas:
        texto = linha.inner_text()
        
        # Verifica se é o relatório correto
        if nome_relatorio.lower() not in texto.lower():
            continue
        
        # Verifica status
        if "Concluído" not in texto:
            continue
        
        # Filtro por GrupoServico
        if filtro_grupo_servico == "vazio":
            if "&GrupoServico=BT" in texto:
                print(f"   ⏭️ Ignorando (tem GrupoServico=BT)")
                continue
        elif filtro_grupo_servico == "BT":
            if "&GrupoServico=BT" not in texto:
                print(f"   ⏭️ Ignorando (não tem GrupoServico=BT)")
                continue
        
        # Extrai data
        m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
        if m:
            dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
            if maior_dt is None or dt > maior_dt:
                maior_dt = dt
                linha_alvo = linha
                # Captura info para debug
                params_start = texto.find("Parametros")
                if params_start != -1:
                    info_adicional = texto[params_start:params_start+200]
                else:
                    info_adicional = texto[:100]
    
    return linha_alvo, maior_dt, info_adicional

# =================================================================
# 8. FUNÇÃO PARA BAIXAR RELATÓRIO
# =================================================================
def baixar_relatorio(page, linha_alvo, nome_relatorio):
    """Faz o download do relatório"""
    try:
        print(f"   📥 Baixando...")
        with page.expect_download() as download_info:
            linha_alvo.locator("button, a, img").first.click()
        
        download = download_info.value
        temp_path = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.zip"
        download.save_as(temp_path)
        
        with open(temp_path, 'rb') as f:
            conteudo_zip = f.read()
        
        os.remove(temp_path)
        return conteudo_zip
    except Exception as e:
        print(f"   ❌ Erro no download: {e}")
        return None

# =================================================================
# 9. FUNÇÃO PRINCIPAL
# =================================================================
def run(playwright: Playwright, modo="simples") -> None:
    """
    modo: "simples" (3 relatórios) ou "completo" (7 relatórios)
    """
    print(f"\n{'='*60}")
    print(f"🚀 INICIANDO COLETOR - MODO: {modo.upper()}")
    print(f"{'='*60}")
    
    # Seleciona os relatórios conforme o modo
    if modo == "simples":
        relatorios = RELATORIOS_SIMPLES
    else:
        relatorios = RELATORIOS_COMPLETOS
    
    # Diretório para perfil do navegador
    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
        os.makedirs(user_dir, exist_ok=True)
    else:
        user_dir = r"C:\Temp\chrome_profile"
    
    launch_args = ['--no-sandbox'] if is_github_actions() else []
    
    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,
        no_viewport=True,
        slow_mo=500,
    )
    page = context.pages[0]
    page.set_default_timeout(120000)

    try:
        # Login
        print("\n🌐 Fazendo login no portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=60000)
        
        # Espera obrigatória de até 60s para a página estar pronta
        usuario_input = page.get_by_role("textbox", name="Usuário")
        usuario_input.wait_for(state="visible", timeout=60000)
        usuario_input.fill(usuario)
        
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Login").click()
        
        # Aguarda a tela pós-login carregar completamente
        page.wait_for_load_state("networkidle", timeout=60000)
        print("✅ Login realizado")
        
        # Navegar para Relatórios Background
        print("\n📊 Navegando para Relatórios Background...")
        menu_bg = page.get_by_text("Relatórios Background")
        menu_bg.wait_for(state="visible", timeout=60000)
        menu_bg.click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        
        # Aguardar tabela
        print("⏳ Aguardando carregamento da tabela...")
        page.wait_for_selector("tbody tr", timeout=60000)
        time.sleep(3)
        
        linhas = page.locator("tbody tr").all()
        print(f"📋 {len(linhas)} linhas encontradas na tabela")
        
        # Processa cada relatório
        relatorios_baixados = 0
        for nome_chave, conf in relatorios.items():
            nome_busca = conf.get("nome_exibicao", nome_chave)
            print(f"\n🔍 Procurando: {nome_busca}")
            
            # Encontra a linha do relatório
            linha_alvo, data_alvo, info = encontrar_linha_relatorio(
                linhas, 
                nome_busca,
                conf.get("filtro_grupo_servico")
            )
            
            if linha_alvo:
                print(f"   ✅ Encontrado (Data: {data_alvo})")
                
                # Faz o download
                conteudo_zip = baixar_relatorio(page, linha_alvo, nome_busca)
                
                if conteudo_zip:
                    # Define o destino
                    if is_github_actions():
                        destino = conf["pasta_sp"]
                    else:
                        destino = PASTAS_LOCAIS[conf["pasta_local"]]
                    
                    # Processa o arquivo com as regras de nomenclatura
                    sucesso = processar_zip(conteudo_zip, conf, destino)
                    
                    if sucesso:
                        relatorios_baixados += 1
                        print(f"   ✅ {nome_busca} processado com sucesso!")
                    else:
                        print(f"   ❌ Falha ao processar {nome_busca}")
                else:
                    print(f"   ❌ Falha no download de {nome_busca}")
            else:
                print(f"   ⚠️ Nenhum relatório 'Concluído' encontrado para: {nome_busca}")
        
        print("\n" + "="*60)
        print(f"✅ COLETOR FINALIZADO - {relatorios_baixados}/{len(relatorios)} relatórios baixados")
        print("="*60)
        
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
        
        try:
            screenshot_path = f"erro_coletor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            page.screenshot(path=screenshot_path, full_page=True)
            print(f"📸 Screenshot salvo: {screenshot_path}")
        except:
            pass
        raise
    finally:
        context.close()
        
        if is_github_actions():
            user_dir = os.path.join(os.getcwd(), "temp_navegador")
            if os.path.exists(user_dir):
                import shutil
                shutil.rmtree(user_dir)
                print("🗑️ Diretório temporário removido")

# =================================================================
# 10. PONTO DE ENTRADA
# =================================================================
if __name__ == "__main__":
    modo = "simples"
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ["completo", "full", "7"]:
            modo = "completo"
        elif arg in ["simples", "simple", "3"]:
            modo = "simples"
    
    print(f"\n🚀 Iniciando Coletor - Modo: {modo.upper()}")
    print(f"⏰ Horário: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    with sync_playwright() as playwright:
        run(playwright, modo)