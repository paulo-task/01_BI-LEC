import os
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
from PIL import Image

# --- FUNÇÃO DE LOG ---
def salvar_log(mensagem):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    caminho_log = "log_captura_direta.txt"
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(f"[{data_hora}] - {mensagem}\n")
    print(f"[{data_hora}] - {mensagem}")

# --- FUNÇÃO DE RECORTE ---
def recortar_print(caminho_img, x1, y1, x2, y2, nome_final):
    try:
        img = Image.open(caminho_img)
        img_recortada = img.crop((x1, y1, x2, y2))
        caminho_final = os.path.join(os.path.dirname(caminho_img), nome_final)
        img_recortada.save(caminho_final)
        if os.path.exists(caminho_img): os.remove(caminho_img)
        salvar_log(f"Recorte concluído: {nome_final}")
    except Exception as e:
        salvar_log(f"Erro no recorte: {e}")

# --- FUNÇÃO DE CAPTURA POWER BI ---
def capturar_telas():
    prints_gerados = {"PAULISTA": None, "PIRATININGA": None}
    URL_DIRETA = "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/50af2f89-ae57-4503-9423-3d55a9b40778/d2e4bc8486f2794906d4?experience=power-bi"

    with sync_playwright() as p:
        user_data_dir = r"C:\Temp\dados_navegador_print"
        save_path = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\PRINTS ZAP"
        if not os.path.exists(save_path): os.makedirs(save_path)

        salvar_log("Iniciando acesso direto ao Power BI...")
        context = p.chromium.launch_persistent_context(
            user_data_dir, headless=False, args=["--start-maximized"], no_viewport=True
        )
        page = context.pages[0]
        agora = datetime.now().strftime("%Y_%m_%d_%H-%M")

        try:
            page.goto(URL_DIRETA, timeout=90000)
            try:
                btn_conta = page.locator("div[role='button'][data-test-id]").first
                if btn_conta.is_visible(timeout=5000):
                    btn_conta.click()
                    page.get_by_role("button", name="Sim").click()
            except: pass

            salvar_log("Aguardando carregamento (20s)...")
            time.sleep(20)

            page.get_by_role("tab", name="ELF Hora").click()
            time.sleep(5)

            def clicar_radio(nome):
                botao = page.get_by_role("radio", name=nome)
                if not botao.is_visible():
                    botao = page.frame_locator("iframe").first.get_by_role("radio", name=nome)
                botao.wait_for(state="visible", timeout=20000)
                botao.click()

            X1, Y1, X2, Y2 = 450, 110, 1743, 878

            # Paulista
            salvar_log("Filtrando PAULISTA...")
            clicar_radio("PAULISTA")
            time.sleep(20)
            path_pau = os.path.join(save_path, f"temp_pau_{agora}.png")
            page.screenshot(path=path_pau)
            recortar_print(path_pau, X1, Y1, X2, Y2, f"PRINT_PAULI_{agora}.png")
            prints_gerados["PAULISTA"] = os.path.join(save_path, f"PRINT_PAULI_{agora}.png")

            # Piratininga
            salvar_log("Filtrando PIRATININGA...")
            clicar_radio("PIRATININGA")
            time.sleep(20)
            path_pira = os.path.join(save_path, f"temp_pira_{agora}.png")
            page.screenshot(path=path_pira)
            recortar_print(path_pira, X1, Y1, X2, Y2, f"PRINT_PIRAT_{agora}.png")
            prints_gerados["PIRATININGA"] = os.path.join(save_path, f"PRINT_PIRAT_{agora}.png")

            context.close()
        except Exception as e:
            salvar_log(f"ERRO POWER BI: {e}")
    return prints_gerados

# --- FUNÇÃO AUXILIAR: ABRE O GRUPO (WHATSAPP WEB) ---
def abrir_grupo(page, grupo):
    salvar_log(f"Abrindo grupo: {grupo}")
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
        
        search_box.wait_for(state="visible", timeout=15000)
        search_box.click()
        time.sleep(0.5)
        
        # Limpa e digita o nome do grupo
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        time.sleep(0.3)
        search_box.fill(grupo)
        time.sleep(3)
        
        # Clica no grupo encontrado nos resultados
        page.get_by_text(grupo, exact=False).first.click()
        time.sleep(3)
        
        page.get_by_test_id("conversation-compose-box-input").wait_for(state="visible", timeout=10000)
        salvar_log(f"✅ Chat carregado: {grupo}")
        return True
    except Exception as e:
        salvar_log(f"❌ Erro ao abrir {grupo}: {e}")
        return False

# --- FUNÇÃO DE ENVIO WHATSAPP (WHATSAPP WEB) ---
def enviar_para_grupos(dicionario_prints):
    regras = [
        {"arquivo": dicionario_prints["PAULISTA"],    "grupos": ["Gestão CPFL Paulista"]},
        {"arquivo": dicionario_prints["PIRATININGA"], "grupos": ["Gestão CPFL Piratininga", "Informativos Administrativo Sorocaba"]}
    ]

    with sync_playwright() as p:
        user_data_zap = r"C:\Temp\dados_navegador_zap"
        context = p.chromium.launch_persistent_context(
            user_data_zap, headless=False, args=["--start-maximized"], no_viewport=True
        )
        page = context.pages[0]
        page.goto("https://web.whatsapp.com")

        try:
            page.get_by_role("textbox", name="Pesquisar ou começar uma nova").wait_for(state="visible", timeout=90000)
            salvar_log("WhatsApp carregado.")

            for regra in regras:
                arquivo = regra["arquivo"]
                if not arquivo or not os.path.exists(arquivo): continue

                for grupo in regra["grupos"]:
                    if not abrir_grupo(page, grupo): continue

                    try:
                        time.sleep(2)
                        
                        # Clique em Anexar
                        btn_anexar = page.get_by_role("button", name="Anexar")
                        btn_anexar.wait_for(state="visible", timeout=10000)
                        btn_anexar.click()
                        time.sleep(1)

                        # Seleção do arquivo via Fotos e vídeos
                        with page.expect_file_chooser() as fc_info:
                            page.get_by_role("menuitem", name="Fotos e vídeos").click()
                        fc_info.value.set_files(arquivo)
                        salvar_log("Arquivo anexado.")
                        time.sleep(2)

                        # Envia via tecla Enter (mais robusto que clicar na seta de enviar)
                        page.keyboard.press("Enter")
                        
                        salvar_log(f"✅ Enviado com sucesso para: {grupo}")
                        time.sleep(8) 
                        
                    except Exception as e:
                        salvar_log(f"⚠️ Falha no envio em {grupo}: {e}")
                        # Fecha qualquer tela aberta (pré-visualização, menus, etc.)
                        for _ in range(5):
                            page.keyboard.press("Escape")
                            time.sleep(0.5)
                        # Aguarda a tela principal do WhatsApp voltar
                        try:
                            page.get_by_role("textbox", name="Pesquisar ou começar uma nova").wait_for(state="visible", timeout=10000)
                            salvar_log("✅ Tela principal recuperada, continuando...")
                        except:
                            salvar_log("⚠️ Não conseguiu voltar à tela principal, tentando próximo grupo mesmo assim...")

            context.close()
        except Exception as e:
            salvar_log(f"ERRO CRÍTICO NO WHATSAPP: {e}")

if __name__ == "__main__":
    salvar_log("=== INÍCIO DA OPERAÇÃO ===")
    prints = capturar_telas()
    enviar_para_grupos(prints)
    salvar_log("=== FIM DA OPERAÇÃO ===\n")