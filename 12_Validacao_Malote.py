import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

IS_GITHUB = os.getenv("GITHUB_ACTIONS") == "true"
SCRIPT_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = SCRIPT_DIR.parent.parent

if IS_GITHUB:
    OUTPUT_DIR = Path(os.getcwd()) / "downloads"
else:
    OUTPUT_DIR = WORKSPACE_DIR / "BI_LEC" / "12_Validacao_Malote"

LOGIN_URL = "https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/login"
MALOTE_URL = "https://cwsilecprd.cpfl.com.br:8443/cwsilecportal/view/malote"
TEMPO_ESPERA_RELATORIO_MS = 120_000


def upload_to_sharepoint(caminho_arquivo: Path, nome_arquivo: str, pasta_sharepoint: str) -> bool:
    """Envia arquivo para o SharePoint via Microsoft Graph API com suporte a arquivos grandes."""
    try:
        sp_client_id = os.getenv("SP_CLIENT_ID", "").strip()
        sp_client_secret = os.getenv("SP_CLIENT_SECRET", "").strip()
        sp_tenant_id = os.getenv("SP_TENANT_ID", "").strip()
        site_url = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"

        if not sp_client_id or not sp_client_secret or not sp_tenant_id:
            print("   ⚠️ SharePoint: credenciais não configuradas para upload.")
            return False

        url_token = f"https://login.microsoftonline.com/{sp_tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": sp_client_id,
            "client_secret": sp_client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url_token, data=data)
        r.raise_for_status()
        token = r.json()["access_token"]
        headers = {"Authorization": f"Bearer {token}"}

        parsed = urlparse(site_url)
        url_site = f"https://graph.microsoft.com/v1.0/sites/{parsed.netloc}:/{parsed.path.strip('/')}"
        r = requests.get(url_site, headers=headers)
        r.raise_for_status()
        site_id = r.json()["id"]

        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        r = requests.get(url_drives, headers=headers)
        r.raise_for_status()

        drive_id = next(
            d["id"]
            for d in r.json().get("value", [])
            if d.get("name") == "Workspace"
        )

        tamanho_total = caminho_arquivo.stat().st_size
        print(f"   ☁️ Iniciando upload para SharePoint ({tamanho_total / (1024 * 1024):.2f} MB): {nome_arquivo}")

        if tamanho_total <= 4 * 1024 * 1024:
            url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content"
            with open(caminho_arquivo, "rb") as f:
                r = requests.put(
                    url_upload,
                    headers={**headers, "Content-Type": "application/octet-stream"},
                    data=f.read(),
                )
            r.raise_for_status()
        else:
            url_session = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/createUploadSession"
            session_payload = {
                "item": {
                    "@microsoft.graph.conflictBehavior": "replace",
                    "name": nome_arquivo,
                }
            }
            r_session = requests.post(url_session, headers=headers, json=session_payload)
            r_session.raise_for_status()
            upload_url = r_session.json()["uploadUrl"]

            chunk_size = 10 * 1024 * 1024
            with open(caminho_arquivo, "rb") as f:
                inicio = 0
                while inicio < tamanho_total:
                    chunk = f.read(chunk_size)
                    fim = inicio + len(chunk) - 1
                    chunk_headers = {
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {inicio}-{fim}/{tamanho_total}",
                    }
                    r_chunk = requests.put(upload_url, headers=chunk_headers, data=chunk)
                    r_chunk.raise_for_status()
                    inicio = fim + 1
                    print(f"      Progresso: {min(100.0, (inicio / tamanho_total) * 100):.1f}%")

        print(f"   ☁️ Upload SharePoint concluído com sucesso: {nome_arquivo}")
        return True
    except Exception as e:
        print(f"   ⚠️ Erro upload SharePoint: {e}")
        return False


def run() -> None:
    pass_file = SCRIPT_DIR / ".pass"
    if pass_file.exists():
        load_dotenv(pass_file)
    else:
        load_dotenv()

    usuario = os.getenv("CPFL_USER")
    senha = os.getenv("CPFL_PASS")
    if not usuario or not senha:
        raise RuntimeError("CPFL_USER e CPFL_PASS devem estar definidos nas variáveis de ambiente ou .pass.")

    hoje = datetime.now()
    data_inicio = hoje.replace(day=1).strftime("%d/%m/%Y")
    data_fim = hoje.strftime("%d/%m/%Y")
    arquivo_saida = OUTPUT_DIR / f"Malote_{hoje:%Y_%m}.xlsx"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=IS_GITHUB)
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = context.new_page()

        try:
            page.goto(LOGIN_URL, timeout=60000)
            page.get_by_role("textbox", name="Usuário").wait_for(
                state="visible", timeout=60000
            )
            page.get_by_role("textbox", name="Usuário").fill(usuario)
            page.get_by_role("textbox", name="Senha").fill(senha)
            page.get_by_role("button", name="Login").click()
            page.wait_for_url("**/view/home", timeout=60000)

            page.goto(MALOTE_URL, timeout=60000)
            page.get_by_role("heading", name="Validação Malote").wait_for(
                state="visible", timeout=60000
            )
            page.locator("button.fa-filter").click()

            datas = page.get_by_role("textbox", name="Data dd/mm/yyyy")
            datas.nth(0).wait_for(state="visible", timeout=30000)
            datas.nth(0).fill(data_inicio)
            datas.nth(1).fill(data_fim)
            page.locator("button.fa-search").click()

            page.get_by_role("grid").wait_for(
                state="visible", timeout=TEMPO_ESPERA_RELATORIO_MS
            )
            botao_exportar = page.locator('button[title="Exportar"]')
            botao_exportar.wait_for(
                state="visible", timeout=TEMPO_ESPERA_RELATORIO_MS
            )
            botao_exportar.click(
                trial=True, timeout=TEMPO_ESPERA_RELATORIO_MS
            )
            with page.expect_download(
                timeout=TEMPO_ESPERA_RELATORIO_MS
            ) as download_info:
                botao_exportar.click(timeout=TEMPO_ESPERA_RELATORIO_MS)

            download = download_info.value
            failure = download.failure()
            if failure:
                raise RuntimeError(f"Falha ao exportar a planilha: {failure}")
            download.save_as(str(arquivo_saida))
            if not arquivo_saida.is_file() or arquivo_saida.stat().st_size == 0:
                raise RuntimeError(f"O arquivo exportado está vazio ou não foi criado: {arquivo_saida}")

            print(f"Planilha salva: {arquivo_saida}")
            print(f"Período aplicado: {data_inicio} a {data_fim}")

            # Upload para SharePoint no GitHub Actions
            if IS_GITHUB:
                upload_to_sharepoint(arquivo_saida, arquivo_saida.name, "BI_LEC/12_Validacao_Malote")

            # Logout do portal CPFL
            try:
                print("Deslogando do sistema para liberar a sessão...")
                page.get_by_role("menuitem", name=re.compile("Logout", re.IGNORECASE)).click(force=True, timeout=5000)
                page.wait_for_load_state("networkidle", timeout=5000)
                print("Sessão encerrada com sucesso.")
            except Exception as e:
                print(f"Aviso: Não foi possível deslogar automaticamente: {e}")

        except PlaywrightTimeoutError as exc:
            if "/view/login" in page.url:
                raise RuntimeError(
                    "O login não foi concluído. Verifique as credenciais ou se o portal "
                    "solicitou validação adicional."
                ) from exc
            raise RuntimeError(
                "O portal não concluiu uma etapa dentro do tempo esperado; "
                "nenhuma exportação foi confirmada."
            ) from exc
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run()
