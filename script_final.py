import os
import time
import base64
import requests
import json
from datetime import datetime
from playwright.sync_api import sync_playwright
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

EVOLUTION_API_URL = os.getenv("EVOLUTION_URL", "http://167.234.240.5:8080")
EVOLUTION_API_KEY = os.getenv("EVOLUTION_API_KEY", "ps0lusm8tk4fvnb67rjq")
EVOLUTION_INSTANCE = os.getenv("EVOLUTION_INSTANCE", "engelmig")
POWERBI_USER = os.getenv("POWERBI_USER")
POWERBI_PASS = os.getenv("POWERBI_PASS")

IS_GITHUB = os.getenv("GITHUB_ACTIONS", "false").lower() == "true"
IS_HEADLESS = IS_GITHUB

TEMP_DIR = r"D:\Repository\01_BI-LEC\prints"
USER_DATA_DIR_PBI = r"D:\Repository\01_BI-LEC\dados_pbi"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(USER_DATA_DIR_PBI, exist_ok=True)

# IDs DOS GRUPOS (ENCONTRADOS!)
GRUPOS_WHATSAPP = {
    "Gestão CPFL Paulista _ UEN 175": "120363244173618567@g.us",
    "Gestão CPFL Piratininga": "120363208473865567@g.us",
    "Informativos Administrativo Sorocaba": "120363230510006463@g.us",
}

def salvar_log(mensagem):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    caminho_log = os.path.join(TEMP_DIR, "log_final.txt")
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(f"[{data_hora}] {mensagem}\n")
    print(f"[{data_hora}] {mensagem}")

def recortar_print(caminho_img, x1, y1, x2, y2, nome_final):
    try:
        img = Image.open(caminho_img)
        img_recortada = img.crop((x1, y1, x2, y2))
        caminho_final = os.path.join(TEMP_DIR, nome_final)
        img_recortada.save(caminho_final)
        if os.path.exists(caminho_img):
            os.remove(caminho_img)
        return caminho_final
    except Exception as e:
        salvar_log(f"Erro recorte: {e}")
        return None

def capturar_telas():
    prints_gerados = {"PAULISTA": None, "PIRATININGA": None}
    URL = "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/50af2f89-ae57-4503-9423-3d55a9b40778/d2e4bc8486f2794906d4?experience=power-bi"

    salvar_log("=== CAPTURA POWER BI ===")
    
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_DIR_PBI, 
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--disable-gpu"],
            no_viewport=True,
            slow_mo=2000
        )
        
        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")

        try:
            salvar_log("Navegando...")
            page.goto(URL, timeout=120000, wait_until="domcontentloaded")
            salvar_log("Aguardando 40s...")
            time.sleep(40)

            X1, Y1, X2, Y2 = 450, 110, 1743, 878

            try:
                page.get_by_role("tab", name="ELF Hora").click(timeout=15000)
                time.sleep(5)
            except:
                pass

            def clicar_radio(nome):
                botao = page.get_by_role("radio", name=nome)
                botao.click(timeout=10000)
                salvar_log(f"Clicou: {nome}")
                time.sleep(25)

            # PAULISTA
            try:
                clicar_radio("PAULISTA")
                path = os.path.join(TEMP_DIR, f"temp_pau_{agora}.png")
                page.screenshot(path=path)
                final = recortar_print(path, X1, Y1, X2, Y2, f"PRINT_PAULI_{agora}.png")
                if final:
                    prints_gerados["PAULISTA"] = final
                    salvar_log("✅ PAULISTA OK")
            except Exception as e:
                salvar_log(f"❌ PAULISTA: {e}")

            # PIRATININGA
            try:
                clicar_radio("PIRATININGA")
                path = os.path.join(TEMP_DIR, f"temp_pira_{agora}.png")
                page.screenshot(path=path)
                final = recortar_print(path, X1, Y1, X2, Y2, f"PRINT_PIRAT_{agora}.png")
                if final:
                    prints_gerados["PIRATININGA"] = final
                    salvar_log("✅ PIRATININGA OK")
            except Exception as e:
                salvar_log(f"❌ PIRATININGA: {e}")

            context.close()
            
        except Exception as e:
            salvar_log(f"ERRO: {e}")
            context.close()

    return prints_gerados

def enviar_whatsapp(arquivo, grupo_nome, grupo_id):
    try:
        if not arquivo or not os.path.exists(arquivo):
            salvar_log(f"Arquivo não encontrado: {arquivo}")
            return False

        with open(arquivo, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        url = f"{EVOLUTION_API_URL}/message/sendMedia/{EVOLUTION_INSTANCE}"
        headers = {"apikey": EVOLUTION_API_KEY, "Content-Type": "application/json"}
        payload = {
            "number": grupo_id,
            "mediatype": "image",
            "media": img_b64,
            "caption": f"Relatório {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        }

        response = requests.post(url, headers=headers, json=payload, timeout=30)
        
        if response.status_code in [200, 201]:
            salvar_log(f"✅ Enviado para {grupo_nome}")
            return True
        else:
            salvar_log(f"❌ Erro ao enviar para {grupo_nome}: {response.status_code}")
            return False

    except Exception as e:
        salvar_log(f"❌ Erro: {e}")
        return False

def main():
    salvar_log("╔════════════════════════════════════════╗")
    salvar_log("║   AUTOMAÇÃO COMPLETA - VERSÃO FINAL    ║")
    salvar_log("╚════════════════════════════════════════╝")

    # 1. Capturar prints
    prints = capturar_telas()

    # 2. Enviar para grupos
    salvar_log("\n=== ENVIANDO PARA WHATSAPP ===")
    
    regras = [
        {"arquivo": prints["PAULISTA"], "grupos": ["Gestão CPFL Paulista _ UEN 175"]},
        {"arquivo": prints["PIRATININGA"], "grupos": ["Gestão CPFL Piratininga", "Informativos Administrativo Sorocaba"]}
    ]

    for regra in regras:
        arquivo = regra["arquivo"]
        if not arquivo:
            continue
        
        for grupo_nome in regra["grupos"]:
            grupo_id = GRUPOS_WHATSAPP.get(grupo_nome)
            if grupo_id:
                enviar_whatsapp(arquivo, grupo_nome, grupo_id)

    salvar_log("\n╔════════════════════════════════════════╗")
    salvar_log("║     AUTOMAÇÃO CONCLUÍDA!               ║")
    salvar_log("╚════════════════════════════════════════╝\n")

if __name__ == "__main__":
    main()