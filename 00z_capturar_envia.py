#!/usr/bin/env python3
"""
SCRIPT DE AUTOMAÇÃO COMPLETO
Captura print do Power BI e envia para grupos WhatsApp
Roda localmente ou no GitHub Actions
"""

import os
import sys
import time
import json
import requests
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

IS_GITHUB = os.getenv("GITHUB_ACTIONS", "false").lower() == "true"
IS_HEADLESS = IS_GITHUB

if IS_GITHUB:
    TEMP_DIR = "/tmp/relatorios"
    USER_DATA_PBI = "/tmp/dados_pbi"
    USER_DATA_ZAP = "/tmp/dados_zap"
    LOG_FILE = "/tmp/relatorios/log_automacao.txt"
else:
    TEMP_DIR = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\03 Repository\01_BI-LEC\relatorios"
    USER_DATA_PBI = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\03 Repository\01_BI-LEC\dados_pbi"
    USER_DATA_ZAP = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\03 Repository\01_BI-LEC\dados_zap"
    LOG_FILE = os.path.join(TEMP_DIR, "log_automacao.txt")

for d in [TEMP_DIR, USER_DATA_PBI, USER_DATA_ZAP]:
    Path(d).mkdir(parents=True, exist_ok=True)

POWERBI_USER = os.getenv("PB_USER")
POWERBI_PASS = os.getenv("PB_PASS")

# Evolution API (WhatsApp) - usado no GitHub Actions
EVOLUTION_URL = os.getenv("EVOLUTION_URL", "")
EVOLUTION_API_KEY = os.getenv("EVOLUTION_API_KEY", "")
EVOLUTION_INSTANCE = os.getenv("EVOLUTION_INSTANCE", "engelmig")

URL_POWERBI = "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/50af2f89-ae57-4503-9423-3d55a9b40778/d2e4bc8486f2794906d4?experience=power-bi"

def log(mensagem):
    data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    msg = f"[{data}] {mensagem}"
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

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

def capturar_powerbi():
    log("=== INICIANDO CAPTURA POWER BI ===")
    prints = {"PAULISTA": None, "PIRATININGA": None}
    
    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox"
        ]
        
        context = p.chromium.launch_persistent_context(
            USER_DATA_PBI,
            headless=IS_HEADLESS,
            args=["--start-maximized"] + args,
            no_viewport=True,
            slow_mo=2000
        )
        
        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")
        
        try:
            log("Acessando Power BI...")
            page.goto(URL_POWERBI, timeout=120000, wait_until="domcontentloaded")
            
            try:
                email_input = page.locator("input[type='email']")
                if email_input.is_visible(timeout=5000):
                    log("Fazendo login...")
                    email_input.fill(POWERBI_USER)
                    page.get_by_role("button", name="Próximo").click()
                    time.sleep(3)
                    pwd_input = page.locator("input[type='password']")
                    pwd_input.fill(POWERBI_PASS)
                    page.get_by_role("button", name="Entrar").click()
                    time.sleep(10)
            except:
                log("Sessão já autenticada")
            
            log("Aguardando carregamento (40s)...")
            time.sleep(40)
            
            try:
                page.get_by_role("tab", name="ELF Hora").click(timeout=15000)
                time.sleep(5)
            except:
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
                                radio.click();
                                break;
                            }
                        }
                    }
                """)
                time.sleep(25)
                path_temp = os.path.join(TEMP_DIR, f"temp_pau_{agora}.png")
                page.screenshot(path=path_temp)
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
                                radio.click();
                                break;
                            }
                        }
                    }
                """)
                time.sleep(25)
                path_temp = os.path.join(TEMP_DIR, f"temp_pira_{agora}.png")
                page.screenshot(path=path_temp)
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

def enviar_via_evolution(arquivo, numero_ou_grupo, is_grupo=True):
    """Envia imagem via Evolution API (usado no GitHub Actions)."""
    try:
        endpoint = f"{EVOLUTION_URL}/message/sendMedia/{EVOLUTION_INSTANCE}"
        headers = {"apikey": EVOLUTION_API_KEY, "Content-Type": "application/json"}
        
        import base64
        with open(arquivo, "rb") as f:
            imagem_b64 = base64.b64encode(f.read()).decode("utf-8")
        
        payload = {
            "number": numero_ou_grupo,
            "mediatype": "image",
            "mimetype": "image/png",
            "caption": f"Relatório Power BI - {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "media": imagem_b64,
            "fileName": os.path.basename(arquivo)
        }
        resp = requests.post(endpoint, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        return True
    except Exception as e:
        log(f"❌ Erro Evolution API: {e}")
        return False


def enviar_whatsapp(prints):
    log("\n=== ENVIANDO PARA WHATSAPP ===")
    regras = [
        {"arquivo": prints["PAULISTA"], "grupos": ["Gestão CPFL Paulista _ UEN 175"]},
        {"arquivo": prints["PIRATININGA"], "grupos": ["Gestão CPFL Piratininga", "Informativos Administrativo Sorocaba"]}
    ]

    if IS_GITHUB:
        # No GitHub Actions: usa Evolution API (não há sessão do WhatsApp Web)
        log("Modo GitHub Actions: usando Evolution API para envio...")
        GRUPOS_EVOLUTION = {
            "Gestão CPFL Paulista _ UEN 175":         os.getenv("GRUPO_PAULISTA_ID", ""),
            "Gestão CPFL Piratininga":                os.getenv("GRUPO_PIRATININGA_ID", ""),
            "Informativos Administrativo Sorocaba":   os.getenv("GRUPO_SOROCABA_ID", ""),
        }
        for regra in regras:
            arquivo = regra["arquivo"]
            if not arquivo or not os.path.exists(arquivo):
                log(f"⚠️ Arquivo não encontrado, pulando: {arquivo}")
                continue
            for grupo_nome in regra["grupos"]:
                grupo_id = GRUPOS_EVOLUTION.get(grupo_nome, "")
                if not grupo_id:
                    log(f"⚠️ ID do grupo não configurado: {grupo_nome}")
                    continue
                if enviar_via_evolution(arquivo, grupo_id):
                    log(f"✅ Enviado para: {grupo_nome}")
                else:
                    log(f"❌ Falha ao enviar para: {grupo_nome}")
        return

    # Execução local: usa WhatsApp Web via Playwright
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_ZAP,
            headless=False,
            args=["--start-maximized"],
            no_viewport=True,
            slow_mo=1000
        )
        page = context.pages[0]
        
        try:
            log("Abrindo WhatsApp Web...")
            page.goto("https://web.whatsapp.com", timeout=60000)
            log("Aguardando carregamento (20s)...")
            time.sleep(20)
            page.wait_for_selector("#pane-side", timeout=60000)
            log("✅ WhatsApp carregado")
            
            for regra in regras:
                arquivo = regra["arquivo"]
                if not arquivo or not os.path.exists(arquivo):
                    continue
                for grupo_nome in regra["grupos"]:
                    if enviar_para_grupo(page, arquivo, grupo_nome):
                        log(f"✅ Enviado para: {grupo_nome}")
                    else:
                        log(f"❌ Falha ao enviar para: {grupo_nome}")
                    time.sleep(5)
            context.close()
        except Exception as e:
            log(f"❌ ERRO: {e}")
            context.close()

def enviar_para_grupo(page, arquivo, grupo_nome):
    try:
        log(f"Procurando grupo: {grupo_nome}")
        grupo = page.locator("#pane-side").get_by_text(grupo_nome, exact=False).first
        if not grupo.is_visible(timeout=5000):
            log(f"Grupo não encontrado: {grupo_nome}")
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
        btn_enviar = page.locator("span[data-icon='send'], span[data-icon='wds-ic-send-filled']").first
        btn_enviar.wait_for(state="visible", timeout=15000)
        btn_enviar.click()
        return True
    except Exception as e:
        log(f"Erro ao enviar para {grupo_nome}: {e}")
        return False

def main():
    log("╔════════════════════════════════════════╗")
    log("║  AUTOMAÇÃO COMPLETA - POWER BI + ZAPP  ║")
    log(f"║  Ambiente: {'GitHub Actions' if IS_GITHUB else 'Local'}              ║")
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