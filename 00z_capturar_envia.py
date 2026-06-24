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
    "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/ff3f2d1f-9433-4060-b072-07b666de8da0/9592b20a8d6c05c3d407?experience=power-bi"
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

    zip_tmp = os.path.join(TEMP_DIR, "dados_zap_dec.zip")
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

def _url_eh_login(url):
    return any(h in url for h in ("login.microsoftonline.com", "login.live.com", "login.windows.net"))


def tentar_login_powerbi(page):
    """Preenche credenciais Microsoft — só na tela de login OAuth."""
    if not _url_eh_login(page.url):
        return False
    if not POWERBI_USER or not POWERBI_PASS:
        log("AVISO: PB_USER/PB_PASS não definidos — login automático indisponível.")
        return False

    try:
        # Seletores exclusivos da Microsoft — nunca usar textbox genérico (pega "Pesquisar" do Power BI)
        email_input = page.locator("#i0116, input[name='loginfmt']").first
        if not email_input.is_visible(timeout=8000):
            btn_conta = page.locator("div[role='button'][data-test-id]").first
            if btn_conta.is_visible(timeout=3000):
                log("Clicando na conta salva...")
                btn_conta.click()
                try:
                    page.get_by_role("button", name="Sim").click(timeout=5000)
                except Exception:
                    pass
            return True

        log("Fazendo login no Power BI (inserindo email)...")
        email_input.fill(POWERBI_USER)
        try:
            page.locator("#idSIButton9, input[type='submit']").first.click(timeout=3000)
        except Exception:
            email_input.press("Enter")
        time.sleep(5)

        email_input_ms = page.locator("#i0116, input[name='loginfmt']").first
        if email_input_ms.is_visible(timeout=5000):
            log("Preenchendo email na tela da Microsoft...")
            email_input_ms.fill(POWERBI_USER)
            try:
                page.locator("#idSIButton9, input[type='submit']").first.click(timeout=3000)
            except Exception:
                email_input_ms.press("Enter")
            time.sleep(3)

        pass_input = page.locator("#i0118, input[name='passwd'], input[type='password']").first
        if pass_input.is_visible(timeout=15000):
            log("Preenchendo senha...")
            pass_input.fill(POWERBI_PASS)
            time.sleep(1)
            try:
                page.locator("#idSIButton9, input[type='submit']").first.click(timeout=3000)
            except Exception:
                pass_input.press("Enter")

        try:
            btn_sim = page.locator("#idSIButton9, input[type='submit']").first
            if btn_sim.is_visible(timeout=10000):
                log("Clicando em Sim para manter conectado...")
                btn_sim.click()
        except Exception:
            pass
        return True
    except Exception as e:
        log(f"Login automático falhou: {e}")
        return False


def limpar_campo_pesquisa(page):
    """Fecha barra de pesquisa do Power BI caso tenha foco ou texto residual."""
    for _ in range(3):
        page.keyboard.press("Escape")
        time.sleep(0.3)
    try:
        page.get_by_role("tab", name="ELF Hora").click(timeout=3000)
    except Exception:
        pass


def aguardar_powerbi_pronto(page, timeout_s=180):
    """Aguarda sair do OAuth e o relatório ficar interativo."""
    log("Aguardando relatório Power BI carregar...")
    inicio = time.time()
    tentativas_login = 0
    while time.time() - inicio < timeout_s:
        if page.is_closed():
            raise RuntimeError("Navegador fechado durante o carregamento do Power BI")

        url = page.url
        if _url_eh_login(url):
            if tentativas_login < 2:
                log("Sessão expirada — tentando login automático...")
                tentar_login_powerbi(page)
                tentativas_login += 1
            else:
                log(f"Ainda em login: {url[:80]}...")
            time.sleep(5)
            continue

        if "app.powerbi.com" in url:
            try:
                page.get_by_role("tab", name="ELF Hora").wait_for(state="visible", timeout=15000)
                log("Relatório Power BI pronto.")
                return
            except Exception:
                pass
        time.sleep(3)

    raise TimeoutError(f"Power BI não carregou em {timeout_s}s (URL: {page.url})")


def clicar_radio(page, nome):
    for tentativa in range(1, 4):
        alvos = [page] + [f for f in page.frames if f != page.main_frame]
        for alvo in alvos:
            try:
                botao = alvo.get_by_role("radio", name=nome)
                if botao.is_visible(timeout=3000):
                    botao.click()
                    return
            except Exception:
                continue
        log(f"Radio '{nome}' não encontrado (tentativa {tentativa}/3)...")
        time.sleep(3)
    raise TimeoutError(f"Radio '{nome}' não encontrado após 3 tentativas")


def aguardar_visual_atualizado(page, segundos=20):
    time.sleep(3)
    try:
        page.wait_for_load_state("networkidle", timeout=45000)
    except Exception:
        pass
    time.sleep(segundos)


def tirar_screenshot_seguro(page, path, tentativas=3):
    for n in range(1, tentativas + 1):
        try:
            if page.is_closed():
                raise RuntimeError("Navegador foi fechado antes do screenshot")
            page.screenshot(path=path, full_page=False, timeout=60000)
            if os.path.getsize(path) > 5000:
                return
            log(f"Screenshot muito pequeno (tentativa {n}/{tentativas}), repetindo...")
        except Exception as e:
            log(f"Screenshot tentativa {n}/{tentativas} falhou: {e}")
            if n == tentativas:
                raise
        time.sleep(5)


def capturar_powerbi():
    log("=== INICIANDO CAPTURA POWER BI ===")
    prints = {"PAULISTA": None, "PIRATININGA": None}

    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--window-size=1920,1080",
        ]
        if not IS_GITHUB:
            args.append("--start-maximized")

        context = p.chromium.launch_persistent_context(
            USER_DATA_PBI,
            headless=IS_HEADLESS,
            args=args,
            viewport={"width": 1920, "height": 1080},
            slow_mo=500 if IS_GITHUB else 200,
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
        )

        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")

        try:
            log("Acessando Power BI...")
            page.goto(URL_POWERBI, timeout=120000)
            log(f"URL após goto: {page.url}")
            time.sleep(5)

            aguardar_powerbi_pronto(page)

            page.get_by_role("tab", name="ELF Hora").click(timeout=60000)
            time.sleep(5)
            limpar_campo_pesquisa(page)

            # Coordenadas exatas medidas pelo usuário na imagem Full HD (1920x1080)
            X1, Y1, X2, Y2 = 260, 85, 1880, 1055

            # PAULISTA
            try:
                log("Capturando PAULISTA...")
                limpar_campo_pesquisa(page)
                clicar_radio(page, "PAULISTA")
                aguardar_visual_atualizado(page)
                
                path_temp = os.path.join(TEMP_DIR, f"PRINT_FULL_pau_{agora}.png")
                tirar_screenshot_seguro(page, path_temp)
                final = recortar_imagem(path_temp, X1, Y1, X2, Y2, f"PRINT_PAULI_{agora}.png")
                if final:
                    prints["PAULISTA"] = final
                    log("✅ PAULISTA OK")
            except Exception as e:
                log(f"❌ Erro PAULISTA: {e}")

            # PIRATININGA
            try:
                log("Capturando PIRATININGA...")
                limpar_campo_pesquisa(page)
                clicar_radio(page, "PIRATININGA")
                aguardar_visual_atualizado(page)
                
                path_temp = os.path.join(TEMP_DIR, f"PRINT_FULL_pira_{agora}.png")
                tirar_screenshot_seguro(page, path_temp)
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

def whatsapp_pediu_qr(page):
    seletores = ["canvas[aria-label*='QR']", "[data-testid='qrcode']"]
    for sel in seletores:
        try:
            if page.locator(sel).first.is_visible(timeout=1500):
                return True
        except Exception:
            pass
    return False


def aguardar_whatsapp_pronto(page, timeout_ms=120000):
    inicio = time.time()
    while (time.time() - inicio) * 1000 < timeout_ms:
        if whatsapp_pediu_qr(page):
            raise RuntimeError(
                "Sessão WhatsApp expirada — QR Code detectado. "
                "No PC, rode: python 00z_gerar_sessao.py, escaneie o QR, "
                "commite whatsapp_session.enc e atualize o Secret WHATSAPP_KEY."
            )
        try:
            if page.locator("#pane-side").first.is_visible(timeout=2000):
                break
        except Exception:
            pass
        try:
            busca = page.get_by_role("textbox", name="Pesquisar ou começar uma nova")
            if busca.first.is_visible(timeout=2000):
                break
        except Exception:
            pass
        time.sleep(2)
    else:
        raise TimeoutError(f"WhatsApp não carregou em {timeout_ms}ms")

    aguardar_sincronizacao_criptografia(page)


def aguardar_sincronizacao_criptografia(page):
    """
    Evita 'Aguardando mensagem' — ocorre quando a mídia é enviada
    antes das chaves de criptografia terminarem de sincronizar.
    """
    log("Aguardando sincronização de criptografia do WhatsApp...")
    try:
        page.wait_for_load_state("networkidle", timeout=90000)
    except Exception:
        pass

    segundos = 60 if IS_GITHUB else 25
    log(f"Estabilizando sessão ({segundos}s)...")
    time.sleep(segundos)


def fechar_dialogos_whatsapp(page):
    """Fecha avisos comuns que bloqueiam o envio."""
    for texto in ("OK", "Continuar", "Entendi"):
        try:
            btn = page.get_by_role("button", name=texto)
            if btn.first.is_visible(timeout=1000):
                btn.first.click()
                time.sleep(0.5)
        except Exception:
            pass


def confirmar_envio_midia(page):
    """Aguarda a pré-visualização carregar e clica em Enviar (mais confiável que Enter)."""
    time.sleep(3)

    seletores_enviar = [
        "span[data-icon='send']",
        "button[aria-label='Enviar']",
        "button[aria-label='Send']",
    ]
    for sel in seletores_enviar:
        try:
            btn = page.locator(sel).last
            if btn.is_visible(timeout=8000):
                btn.click()
                log("Botão Enviar clicado na pré-visualização.")
                break
        except Exception:
            continue
    else:
        page.keyboard.press("Enter")
        log("Enviado via Enter (fallback).")

    # Tempo extra para upload e criptografia da mídia
    espera = 20 if IS_GITHUB else 12
    log(f"Aguardando confirmação de envio ({espera}s)...")
    time.sleep(espera)


def enviar_whatsapp(prints):
    log("\n=== ENVIANDO PARA WHATSAPP ===")

    if IS_GITHUB:
        ok = restaurar_sessao_zap()
        if not ok:
            log("❌ Sem sessão WhatsApp — envio cancelado.")
            log("   → No PC: python 00z_gerar_sessao.py")
            log("   → Commite whatsapp_session.enc no repositório")
            log("   → Adicione whatsapp_key.txt como Secret WHATSAPP_KEY no GitHub")
            return

    regras = [
        {"arquivo": prints["PAULISTA"],    "grupos": ["Gestão CPFL Paulista _ UEN 175"]},
        {"arquivo": prints["PIRATININGA"], "grupos": ["Gestão CPFL Piratininga", "Informativos Administrativo Sorocaba"]},
    ]

    with sync_playwright() as p:
        args_zap = [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--window-size=1920,1080",
        ]

        context = p.chromium.launch_persistent_context(
            USER_DATA_ZAP,
            headless=False,
            args=args_zap,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            slow_mo=500,
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
        )
        page = context.pages[0]
        page.set_default_timeout(120000)

        try:
            log("Abrindo WhatsApp Web...")
            page.goto("https://web.whatsapp.com", timeout=120000, wait_until="domcontentloaded")
            log("Aguardando carregamento da tela principal...")
            aguardar_whatsapp_pronto(page, timeout_ms=120000)
            fechar_dialogos_whatsapp(page)
            log("✅ WhatsApp carregado e sincronizado")

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
                    time.sleep(15 if IS_GITHUB else 8)

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
        except Exception:
            pass

        time.sleep(2)
        log(f"✅ Chat carregado: {grupo_nome}")
        
        # Clique em Anexar
        btn_anexar = page.get_by_role("button", name="Anexar")
        if not btn_anexar.is_visible(timeout=5000):
             btn_anexar = page.locator("span[data-icon='clip'], span[data-icon='plus']").first
             
        btn_anexar.wait_for(state="visible", timeout=10000)
        btn_anexar.click()
        time.sleep(1)

        # Fotos e vídeos — aguarda pré-visualização antes de enviar
        with page.expect_file_chooser() as fc_info:
            opcao = page.get_by_role("menuitem", name="Fotos e vídeos")
            if not opcao.is_visible(timeout=3000):
                 opcao = page.locator("li:has-text('Fotos e vídeos'), li:has-text('Galeria')").first
            opcao.click()
            
        fc_info.value.set_files(arquivo)
        log(f"Arquivo selecionado: {arquivo}")
        confirmar_envio_midia(page)
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