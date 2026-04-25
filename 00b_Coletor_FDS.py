import os
import time
import re
import sys
import io
import zipfile
import shutil
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

print("✅ Credenciais CPFL carregadas")

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
# 3. RELATÓRIOS DO FDS (PDL + Não Visitadas)
# =================================================================
def get_nome_arquivo(conf, ext):
    agora = datetime.now()
    ano_mes_dia = agora.strftime('%Y_%m_%d')
    return f"{conf['nome_base']}_{ano_mes_dia}{ext}"

RELATORIOS_FDS = {
    "Produtividade Diária Leiturista - Analítico": {
        "pasta_sp": "BI_LEC/02_PDL_Analitico",
        "pasta_local": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
        "nome_base": "Produtividade Diária Leiturista - Analítico",
        "filtro_grupo_servico": None,
        "nome_exibicao": "Produtividade Diária Leiturista - Analítico"
    },
    "Instalações Não Visitadas": {
        "pasta_sp": "BI_LEC/04_N_Visitado_Diario",
        "pasta_local": r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario",
        "nome_base": "Instalações Não Visitadas",
        "filtro_grupo_servico": "vazio",
        "nome_exibicao": "Instalações Não Visitadas"
    }
}

# =================================================================
# 4. UPLOAD SHAREPOINT
# =================================================================
def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    try:
        import requests
        from urllib.parse import urlparse

        SP_CLIENT_ID     = os.getenv("SP_CLIENT_ID", "").strip()
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "").strip()
        SP_TENANT_ID     = os.getenv("SP_TENANT_ID", "").strip()
        SITE_URL         = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"

        if not SP_CLIENT_ID:
            print("   ⚠️ SharePoint: credenciais não configuradas")
            return False

        # Token
        r = requests.post(
            f"https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token",
            data={
                "grant_type": "client_credentials",
                "client_id": SP_CLIENT_ID,
                "client_secret": SP_CLIENT_SECRET,
                "scope": "https://graph.microsoft.com/.default"
            }
        )
        r.raise_for_status()
        token = r.json()["access_token"]

        # Site ID
        parsed = urlparse(SITE_URL)
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{parsed.netloc}:/{parsed.path.strip('/')}",
            headers={"Authorization": f"Bearer {token}"}
        )
        r.raise_for_status()
        site_id = r.json()["id"]

        # Drive Workspace
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
            headers={"Authorization": f"Bearer {token}"}
        )
        r.raise_for_status()
        drive_id = next(
            (d["id"] for d in r.json().get("value", []) if d.get("name") == "Workspace"),
            None
        )
        if not drive_id:
            raise Exception("Drive 'Workspace' não encontrado")

        # Upload
        r = requests.put(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"},
            data=conteudo_bytes
        )
        r.raise_for_status()
        print(f"   ☁️ Upload SharePoint: {nome_arquivo}")
        return True

    except Exception as e:
        print(f"   ⚠️ Erro upload SharePoint: {e}")
        return False

# =================================================================
# 5. SALVAR LOCALMENTE
# =================================================================
def salvar_localmente(conteudo_bytes, nome_arquivo, pasta_destino):
    try:
        os.makedirs(pasta_destino, exist_ok=True)
        caminho = os.path.join(pasta_destino, nome_arquivo)
        with open(caminho, "wb") as f:
            f.write(conteudo_bytes)
        print(f"   💾 Salvo localmente: {caminho}")
        return True
    except Exception as e:
        print(f"   ❌ Erro ao salvar localmente: {e}")
        return False

# =================================================================
# 6. PROCESSAR ZIP
# =================================================================
def processar_zip(conteudo_zip, conf, destino):
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo_zip)) as z:
            arquivo_interno = z.namelist()[0]
            ext = os.path.splitext(arquivo_interno)[1]
            nome_final = get_nome_arquivo(conf, ext)
            with z.open(arquivo_interno) as f:
                conteudo = f.read()

        if is_github_actions():
            return upload_to_sharepoint(conteudo, nome_final, destino)
        else:
            return salvar_localmente(conteudo, nome_final, destino)
    except Exception as e:
        print(f"   ❌ Erro processar ZIP: {e}")
        return False

# =================================================================
# 7. ENCONTRAR LINHA DO RELATÓRIO
# =================================================================
def encontrar_linha_relatorio(linhas, nome_relatorio, filtro_grupo_servico=None):
    maior_dt = None
    linha_alvo = None

    for linha in linhas:
        texto = linha.inner_text()

        if nome_relatorio.lower() not in texto.lower():
            continue
        if "Concluído" not in texto:
            continue

        if filtro_grupo_servico == "vazio" and "&GrupoServico=BT" in texto:
            continue
        elif filtro_grupo_servico == "BT" and "&GrupoServico=BT" not in texto:
            continue

        m = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})", texto)
        if m:
            dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M:%S")
            if maior_dt is None or dt > maior_dt:
                maior_dt = dt
                linha_alvo = linha

    return linha_alvo, maior_dt

# =================================================================
# 8. BAIXAR RELATÓRIO
# =================================================================
def baixar_relatorio(page, linha_alvo):
    try:
        print("   📥 Baixando...")
        with page.expect_download() as download_info:
            linha_alvo.locator("button, a, img").first.click()
        download = download_info.value
        temp_path = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.zip"
        download.save_as(temp_path)
        with open(temp_path, "rb") as f:
            conteudo = f.read()
        os.remove(temp_path)
        return conteudo
    except Exception as e:
        print(f"   ❌ Erro no download: {e}")
        return None

# =================================================================
# 9. FUNÇÃO PRINCIPAL
# =================================================================
def run(playwright: Playwright) -> None:
    print(f"\n{'='*60}")
    print("🏖️ INICIANDO COLETOR FDS")
    print(f"{'='*60}")

    if is_github_actions():
        user_dir = os.path.join(os.getcwd(), "temp_navegador")
        os.makedirs(user_dir, exist_ok=True)
    else:
        user_dir = r"C:\Temp\dados_navegador_coletor"

    launch_args = ["--no-sandbox"] if is_github_actions() else []

    context = playwright.chromium.launch_persistent_context(
        user_dir,
        headless=USE_HEADLESS,
        args=launch_args,
        no_viewport=True,
        slow_mo=500,
    )
    page = context.pages[0]
    page.set_default_timeout(180000)

    try:
        print("\n🌐 Fazendo login no portal CPFL...")
        page.goto("https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login", timeout=90000)
        page.wait_for_load_state("networkidle")

        if page.get_by_role("textbox", name="Usuário").is_visible(timeout=15000):
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()
            page.wait_for_load_state("networkidle")
        print("✅ Login realizado")

        print("\n📊 Navegando para Relatórios Background...")
        page.get_by_text("Relatórios Background").click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)

        print("⏳ Aguardando carregamento da tabela...")
        page.wait_for_selector("tbody tr", timeout=120000)
        time.sleep(25)

        linhas = page.locator("tbody tr").all()
        print(f"📋 {len(linhas)} linhas encontradas na tabela")

        baixados = 0
        for nome_chave, conf in RELATORIOS_FDS.items():
            nome_busca = conf["nome_exibicao"]
            print(f"\n🔍 Procurando: {nome_busca}")

            linha_alvo, data_alvo = encontrar_linha_relatorio(
                linhas, nome_busca, conf.get("filtro_grupo_servico")
            )

            if linha_alvo:
                print(f"   ✅ Encontrado (Data: {data_alvo})")
                conteudo_zip = baixar_relatorio(page, linha_alvo)

                if conteudo_zip:
                    destino = conf["pasta_sp"] if is_github_actions() else conf["pasta_local"]
                    if processar_zip(conteudo_zip, conf, destino):
                        baixados += 1
                        print(f"   ✅ {nome_busca} processado com sucesso!")
                    else:
                        print(f"   ❌ Falha ao processar {nome_busca}")
                else:
                    print(f"   ❌ Falha no download de {nome_busca}")
            else:
                print(f"   ⚠️ Nenhum relatório 'Concluído' encontrado para: {nome_busca}")

        print(f"\n{'='*60}")
        print(f"✅ COLETOR FDS FINALIZADO - {baixados}/{len(RELATORIOS_FDS)} relatórios baixados")
        print(f"{'='*60}")

    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
        raise

    finally:
        context.close()
        if is_github_actions():
            user_dir = os.path.join(os.getcwd(), "temp_navegador")
            if os.path.exists(user_dir):
                shutil.rmtree(user_dir)
                print("🗑️ Diretório temporário removido")

# =================================================================
# 10. PONTO DE ENTRADA
# =================================================================
if __name__ == "__main__":
    print(f"\n🏖️ Coletor FDS")
    print(f"⏰ Horário: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with sync_playwright() as playwright:
        run(playwright)