#!/usr/bin/env python3
"""
SCRIPT DE AUTOMAÇÃO COMPLETO
Captura print do Power BI e envia para grupos WhatsApp
Roda localmente ou no GitHub Actions (com sessão WhatsApp restaurada via Secret)
"""

import os
import sys
import time
import zipfile
import shutil
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from PIL import Image
from dotenv import load_dotenv
from cryptography.fernet import Fernet

load_dotenv()

IS_GITHUB = os.getenv("GITHUB_ACTIONS", "false").lower() == "true"
IS_HEADLESS = IS_GITHUB

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if IS_GITHUB:
    TEMP_DIR        = "/tmp/relatorios"
    USER_DATA_PBI   = "/tmp/dados_pbi"
    USER_DATA_ZAP   = "/tmp/dados_zap"
    LOG_FILE        = "/tmp/relatorios/log_automacao.txt"
    SESSION_ENC     = os.path.join(_SCRIPT_DIR, "whatsapp_session.enc")
else:
    # Local - usa o diretório do script
    TEMP_DIR        = os.path.join(_SCRIPT_DIR, "relatorios")
    USER_DATA_PBI   = os.path.join(_SCRIPT_DIR, "dados_pbi")
    USER_DATA_ZAP   = os.path.join(_SCRIPT_DIR, "dados_zap")
    LOG_FILE        = os.path.join(TEMP_DIR, "log_automacao.txt")
    SESSION_ENC     = os.path.join(_SCRIPT_DIR, "whatsapp_session.enc")

for d in [TEMP_DIR, USER_DATA_PBI, USER_DATA_ZAP]:
    Path(d).mkdir(parents=True, exist_ok=True)

POWERBI_USER = os.getenv("PB_USER")
POWERBI_PASS = os.getenv("PB_PASS")
WHATSAPP_KEY = os.getenv("WHATSAPP_KEY", "")

URL_POWERBI = (
    "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b"
    "/reports/50af2f89-ae57-4503-9423-3d55a9b40778"
    "/d2e4bc8486f2794906d4?experience=power-bi"
)

# ─────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────

def log(mensagem):
    data = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    msg = f"[{data}] {mensagem}"
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def restaurar_sessao_zap():
    """Descriptografa whatsapp_session.enc usando WHATSAPP_KEY e restaura em USER_DATA_ZAP."""
    if not WHATSAPP_KEY:
        log("AVISO: WHATSAPP_KEY não definida — WhatsApp sem sessão prévia.")
        return False
    if not os.path.exists(SESSION_ENC):
        log(f"AVISO: Arquivo de sessão não encontrado: {SESSION_ENC}")
        return False

    zip_tmp = "/tmp/dados_zap_dec.zip"
    try:
        f = Fernet(WHATSAPP_KEY.encode())
        with open(SESSION_ENC, "rb") as fp:
            dados_enc = fp.read()
        dados_zip = f.decrypt(dados_enc)
        with open(zip_tmp, "wb") as fp:
            fp.write(dados_zip)

        if os.path.exists(USER_DATA_ZAP):
            shutil.rmtree(USER_DATA_ZAP)
        with zipfile.ZipFile(zip_tmp, "r") as z:
            z.extractall(USER_DATA_ZAP)

        log("✅ Sessão WhatsApp restaurada com sucesso.")
        return True
    except Exception as e:
        log(f"❌ Erro ao restaurar sessão: {e}")
        return False


def recortar_imagem(caminho_img, x1, y1, x2, y2, nome_final):
    try:
        img = Image.open(caminho_img)
        img_recortada = img.crop((x1, y1, x2, y2))
        caminho_final = os.path.join(TEMP_DIR, nome_final)
        img_recortada.save(caminho_final)
        if os.path.exists(caminho_img):
            os.remove(caminho_img)
        return caminho_final
    except Exception as e:
        log(f"Erro recorte: {e}")
        return None


# ─────────────────────────────────────────────
# CAPTURA POWER BI
# ─────────────────────────────────────────────

def capturar_powerbi():
    log("=== INICIANDO CAPTURA POWER BI ===")
    prints = {"PAULISTA": None, "PIRATININGA": None}

    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox",
        ]
        if IS_GITHUB:
            args += ["--window-size=1920,1080"]

        context = p.chromium.launch_persistent_context(
            USER_DATA_PBI,
            headless=IS_HEADLESS,
            args=args,
            viewport={"width": 1920, "height": 1080} if IS_GITHUB else None,
            no_viewport=not IS_GITHUB,
            slow_mo=1000,
        )

        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")

        try:
            log("Acessando Power BI...")
            page.goto(URL_POWERBI, timeout=120000, wait_until="domcontentloaded")

            # Login (só se necessário)
            try:
                email_input = page.locator("input[type='email']")
                if email_input.is_visible(timeout=8000):
                    log("Fazendo login...")
                    email_input.fill(POWERBI_USER)
                    page.get_by_role("button", name="Próximo").click()
                    
                    page.locator("input[type='password']").wait_for(state="visible", timeout=15000)
                    page.locator("input[type='password']").fill(POWERBI_PASS)
                    page.get_by_role("button", name="Entrar").click()
                    
                    # Trata "Continuar conectado?" (Stay signed in?)
                    try:
                        btn_sim = page.get_by_role("button", name="Sim")
                        if btn_sim.is_visible(timeout=10000):
                            btn_sim.click()
                    except:
                        pass
                    
                    log("Aguardando carregamento pós-login (20s)...")
                    time.sleep(20)
            except Exception as e:
                log(f"Aviso de login: {e} (Pode ser que já esteja logado)")

            log("Aguardando carregamento (40s)...")
            time.sleep(40)

            try:
                page.get_by_role("tab", name="ELF Hora").click(timeout=15000)
                time.sleep(5)
            except Exception:
                log("Aviso: Aba não encontrada")

            X1, Y1, X2, Y2 = 450, 110, 1743, 878

            # PAULISTA
            try:
                log("Capturando PAULISTA...")
                page.evaluate("""
                    () => {
                        const radios = document.querySelectorAll('input[type="radio"]');
                        for (let radio of radios) {
                            if (radio.getAttribute('aria-label') && radio.getAttribute('aria-label').includes('PAULISTA')) {
                                radio.click(); break;
                            }
                        }
                    }
                """)
                time.sleep(25)
                path_temp = os.path.join(TEMP_DIR, f"temp_pau_{agora}.png")
                page.screenshot(path=path_temp, full_page=False)
                final = recortar_imagem(path_temp, X1, Y1, X2, Y2, f"PRINT_PAULI_{agora}.png")
                if final:
                    prints["PAULISTA"] = final
                    log("✅ PAULISTA OK")
            except Exception as e:
                log(f"❌ Erro PAULISTA: {e}")

            # PIRATININGA
            try:
                log("Capturando PIRATININGA...")
                page.evaluate("""
                    () => {
                        const radios = document.querySelectorAll('input[type="radio"]');
                        for (let radio of radios) {
                            if (radio.getAttribute('aria-label') && radio.getAttribute('aria-label').includes('PIRATININGA')) {
                                radio.click(); break;
                            }
                        }
                    }
                """)
                time.sleep(25)
                path_temp = os.path.join(TEMP_DIR, f"temp_pira_{agora}.png")
                page.screenshot(path=path_temp, full_page=False)
                final = recortar_imagem(path_temp, X1, Y1, X2, Y2, f"PRINT_PIRAT_{agora}.png")
                if final:
                    prints["PIRATININGA"] = final
                    log("✅ PIRATININGA OK")
            except Exception as e:
                log(f"❌ Erro PIRATININGA: {e}")

            context.close()
        except Exception as e:
            log(f"❌ ERRO CRÍTICO: {e}")
            context.close()

    return prints


# ─────────────────────────────────────────────
# ENVIO WHATSAPP
# ─────────────────────────────────────────────

def enviar_whatsapp(prints):
    log("\n=== ENVIANDO PARA WHATSAPP ===")

    # No GitHub Actions: restaura sessão antes de abrir o browser
    if IS_GITHUB:
        ok = restaurar_sessao_zap()
        if not ok:
            log("❌ Sem sessão WhatsApp — envio cancelado.")
            log("   → Gere o Secret WHATSAPP_SESSION rodando: python 00z_gerar_sessao.py")
            return

    regras = [
        {"arquivo": prints["PAULISTA"],    "grupos": ["Gestão CPFL Paulista _ UEN 175"]},
        {"arquivo": prints["PIRATININGA"], "grupos": ["Gestão CPFL Piratininga", "Informativos Administrativo Sorocaba"]},
    ]

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_ZAP,
            headless=IS_HEADLESS,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080} if IS_HEADLESS else None,
            no_viewport=not IS_HEADLESS,
            slow_mo=1000,
        )
        page = context.pages[0]

        try:
            log("Abrindo WhatsApp Web...")
            page.goto("https://web.whatsapp.com", timeout=60000)
            log("Aguardando carregamento (30s)...")
            time.sleep(30)
            page.wait_for_selector("#pane-side", timeout=60000)
            log("✅ WhatsApp carregado")

            for regra in regras:
                arquivo = regra["arquivo"]
                if not arquivo or not os.path.exists(arquivo):
                    log(f"⚠️  Arquivo não encontrado, pulando")
                    continue
                for grupo_nome in regra["grupos"]:
                    if enviar_para_grupo(page, arquivo, grupo_nome):
                        log(f"✅ Enviado para: {grupo_nome}")
                    else:
                        log(f"❌ Falha ao enviar para: {grupo_nome}")
                    time.sleep(5)

            context.close()
        except Exception as e:
            log(f"❌ ERRO WhatsApp: {e}")
            try:
                erro_path = os.path.join(TEMP_DIR, "PRINT_ERRO_WHATSAPP.png")
                page.screenshot(path=erro_path)
                log(f"📸 Screenshot do erro salvo em: {erro_path}")
            except:
                pass
            context.close()


def enviar_para_grupo(page, arquivo, grupo_nome):
    try:
        log(f"Procurando grupo: {grupo_nome}")
        
        # Clica na caixa de pesquisa
        search_box = page.locator("div[contenteditable='true']").first
        search_box.click()
        # Limpa e digita o nome
        search_box.fill("")
        search_box.fill(grupo_nome)
        time.sleep(3)
        
        # Seleciona o resultado
        grupo = page.locator("#pane-side").get_by_text(grupo_nome, exact=False).first
        if not grupo.is_visible(timeout=8000):
            log(f"Grupo não encontrado na busca: {grupo_nome}")
            return False
            
        grupo.click()
        time.sleep(3)

        btn_anexar = page.get_by_role("button", name="Anexar")
        btn_anexar.wait_for(state="visible", timeout=10000)
        btn_anexar.click()
        time.sleep(1)

        with page.expect_file_chooser() as fc_info:
            page.get_by_text("Fotos e vídeos").click()
        fc_info.value.set_files(arquivo)
        log(f"Arquivo selecionado: {arquivo}")
        time.sleep(3)

        btn_enviar = page.locator(
            "span[data-icon='send'], span[data-icon='wds-ic-send-filled']"
        ).first
        btn_enviar.wait_for(state="visible", timeout=15000)
        btn_enviar.click()
        return True
    except Exception as e:
        log(f"Erro ao enviar para {grupo_nome}: {e}")
        return False


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    log("╔════════════════════════════════════════╗")
    log("║  AUTOMAÇÃO COMPLETA - POWER BI + ZAPP  ║")
    log(f"║  Ambiente: {'GitHub Actions' if IS_GITHUB else 'Local            '}              ║")
    log("╚════════════════════════════════════════╝")

    prints = capturar_powerbi()

    if prints["PAULISTA"] or prints["PIRATININGA"]:
        enviar_whatsapp(prints)
    else:
        log("❌ Nenhum print foi capturado")

    log("\n╔════════════════════════════════════════╗")
    log("║  AUTOMAÇÃO CONCLUÍDA!                  ║")
    log("╚════════════════════════════════════════╝\n")


if __name__ == "__main__":
    main()