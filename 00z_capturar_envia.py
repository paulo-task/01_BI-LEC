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
        # Mantido para debug: a imagem original ficará nos artefatos com o nome PRINT_FULL
        # if os.path.exists(caminho_img):
        #     os.remove(caminho_img)
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
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
        )

        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")

        try:
            log("Acessando Power BI...")
            page.goto(URL_POWERBI, timeout=120000)
            log(f"URL após goto: {page.url}")

            # Login (só se necessário)
            try:
                time.sleep(10) # Aguarda redirects iniciais
                log(f"URL após 10s: {page.url}")
                # 1. Verifica se pediu o email (tela de login nova ou tela de 'Enter email' do Power BI)
                email_input = page.locator("input[type='email'], input[name='loginfmt'], input[placeholder='Enter email'], input[placeholder*='email']").first
                
                # Se não achar pelos seletores acima, tenta pegar o primeiro textbox na tela
                if not email_input.is_visible(timeout=5000):
                    fallback = page.get_by_role("textbox").first
                    if fallback.is_visible(timeout=2000):
                        email_input = fallback

                if email_input.is_visible(timeout=5000):
                    log("Fazendo login no Power BI (inserindo email)...")
                    if POWERBI_USER:
                        email_input.fill(POWERBI_USER)
                    
                    # Tenta clicar em "Enviar" / "Submit" ou aperta Enter
                    try:
                        btn_enviar = page.get_by_role("button", name="Enviar")
                        if btn_enviar.is_visible(timeout=2000):
                            btn_enviar.click()
                        else:
                            email_input.press("Enter")
                    except:
                        email_input.press("Enter")
                    
                    # Aguarda para ver se vai para a tela de senha ou se pede email novamente
                    time.sleep(5)
                    
                    # Pode ser que ele tenha redirecionado para a tela oficial da Microsoft e peça o email de novo
                    email_input_ms = page.locator("input[type='email'], input[name='loginfmt']").first
                    if email_input_ms.is_visible(timeout=5000):
                        log("Preenchendo email novamente na tela da Microsoft...")
                        if POWERBI_USER:
                            email_input_ms.fill(POWERBI_USER)
                        email_input_ms.press("Enter")
                        time.sleep(3)
                    
                    log("Aguardando campo de senha...")
                    pass_input = page.locator("input[type='password'], input[name='passwd']").first
                    pass_input.wait_for(state="visible", timeout=20000)
                    
                    if POWERBI_PASS:
                        pass_input.fill(POWERBI_PASS)
                    time.sleep(1)
                    
                    try:
                        btn_entrar = page.get_by_role("button", name="Entrar")
                        if btn_entrar.is_visible(timeout=2000):
                            btn_entrar.click()
                        else:
                            pass_input.press("Enter")
                    except:
                        pass_input.press("Enter")
                    
                    # Clica no "Sim" (Continuar conectado)
                    try:
                        btn_sim = page.locator("input[type='submit'], button[type='submit'], #idSIButton9").first
                        if btn_sim.is_visible(timeout=10000):
                            log("Clicando em Sim para manter conectado...")
                            btn_sim.click()
                    except:
                        pass
                else:
                    # 2. Verifica se apareceu a tela de escolher conta já salva
                    btn_conta = page.locator("div[role='button'][data-test-id]").first
                    if btn_conta.is_visible(timeout=3000):
                        log("Clicando na conta salva...")
                        btn_conta.click()
                        try:
                            page.get_by_role("button", name="Sim").click(timeout=5000)
                        except:
                            pass
                    
                log("Aguardando carregamento pós-login (20s)...")
                time.sleep(20)
            except Exception as e:
                log(f"Aviso de login: não foi necessário ou algo falhou ({e})")

            log("Aguardando carregamento (20s)...")
            time.sleep(20)

            page.get_by_role("tab", name="ELF Hora").click(timeout=60000)
            time.sleep(5)

            def clicar_radio(nome):
                botao = page.get_by_role("radio", name=nome)
                if not botao.is_visible():
                    botao = page.frame_locator("iframe").first.get_by_role("radio", name=nome)
                botao.wait_for(state="visible", timeout=20000)
                botao.click()

            # Coordenadas exatas medidas pelo usuário na imagem Full HD (1920x1080)
            X1, Y1, X2, Y2 = 260, 85, 1880, 1055

            # PAULISTA
            try:
                log("Capturando PAULISTA...")
                clicar_radio("PAULISTA")
                time.sleep(20)
                
                path_temp = os.path.join(TEMP_DIR, f"PRINT_FULL_pau_{agora}.png")
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
                clicar_radio("PIRATININGA")
                time.sleep(20)
                
                path_temp = os.path.join(TEMP_DIR, f"PRINT_FULL_pira_{agora}.png")
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
            try:
                erro_path = os.path.join(TEMP_DIR, "PRINT_ERRO_CRITICO.png")
                page.screenshot(path=erro_path)
                log(f"📸 Screenshot do erro salvo em: {erro_path}")
            except:
                pass
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
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
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
    log(f"Abrindo grupo: {grupo_nome}")
    try:
        # Fecha qualquer overlay/menu aberto
        for _ in range(3):
            page.keyboard.press("Escape")
            time.sleep(0.3)
        
        # Usa atalho Ctrl+F para garantir foco na barra de busca
        time.sleep(1)
        
        # Tenta localizar a barra de pesquisa
        search_box = page.get_by_role("textbox", name="Pesquisar ou começar uma nova")
        if not search_box.is_visible(timeout=3000):
            # Fallback: clica na área de pesquisa via atalho
            page.keyboard.press("Control+f")
            time.sleep(1)
            search_box = page.get_by_role("textbox", name="Pesquisar ou começar uma nova")
            
        # Se ainda não achar, usa um seletor genérico
        if not search_box.is_visible(timeout=2000):
            search_box = page.locator("div[contenteditable='true'], [role='textbox']").first
        
        search_box.wait_for(state="visible", timeout=15000)
        search_box.click()
        time.sleep(0.5)
        
        # Limpa e digita o nome do grupo
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        time.sleep(0.3)
        search_box.fill(grupo_nome)
        time.sleep(3)
        
        # Clica no grupo encontrado nos resultados
        page.get_by_text(grupo_nome, exact=False).first.click()
        time.sleep(3)
        
        # Tenta achar a caixa de texto de conversa para garantir que abriu
        try:
            page.get_by_test_id("conversation-compose-box-input").wait_for(state="visible", timeout=10000)
        except:
            pass # Ignora se não achar pelo test-id, pois o grupo pode ter aberto mesmo assim
            
        log(f"✅ Chat carregado: {grupo_nome}")
        
        # Clique em Anexar
        btn_anexar = page.get_by_role("button", name="Anexar")
        if not btn_anexar.is_visible(timeout=5000):
             # Fallback para ícone
             btn_anexar = page.locator("span[data-icon='clip'], span[data-icon='plus']").first
             
        btn_anexar.wait_for(state="visible", timeout=10000)
        btn_anexar.click()
        time.sleep(1)

        # Seleção do arquivo via Fotos e vídeos
        with page.expect_file_chooser() as fc_info:
            opcao = page.get_by_role("menuitem", name="Fotos e vídeos")
            if not opcao.is_visible(timeout=3000):
                 opcao = page.locator("li:has-text('Fotos e vídeos'), li:has-text('Galeria')").first
            opcao.click()
            
        fc_info.value.set_files(arquivo)
        log(f"Arquivo selecionado: {arquivo}")
        time.sleep(2)

        # Envia via tecla Enter (mais robusto que clicar na seta de enviar)
        page.keyboard.press("Enter")
        
        time.sleep(8) 
        return True
        
    except Exception as e:
        import traceback
        log(f"⚠️ Falha no envio em {grupo_nome}: {e}")
        log(traceback.format_exc())
        try:
            safe_name = "".join([c for c in grupo_nome if c.isalnum() or c in (' ', '_')]).replace(' ', '_')
            erro_path = os.path.join(TEMP_DIR, f"PRINT_ERRO_ZAP_{safe_name}.png")
            page.screenshot(path=erro_path)
            log(f"📸 Screenshot do erro salvo em: {erro_path}")
        except:
            pass
            
        # Tenta limpar a tela caso falhe
        for _ in range(5):
            page.keyboard.press("Escape")
            time.sleep(0.5)
            
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