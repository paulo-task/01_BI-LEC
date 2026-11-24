#!/usr/bin/env python3
"""
Abre o WhatsApp Web no perfil persistente local (dados_zap).
Use para verificar se o WhatsApp esta conectado ou para escanear o QR Code no seu PC.
"""
import os
import time
from playwright.sync_api import sync_playwright

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_ZAP = os.path.join(SCRIPT_DIR, "dados_zap")

print("=" * 60)
print("  VERIFICANDO / CONECTANDO WHATSAPP WEB NO SEU PC")
print("=" * 60)
print(f"Diretorio da sessao local: {USER_DATA_ZAP}")
print("Abrindo navegador...")

with sync_playwright() as p:
    args = [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-blink-features=AutomationControlled",
        "--disable-gpu",
        "--window-size=1280,900",
    ]
    context = p.chromium.launch_persistent_context(
        USER_DATA_ZAP,
        headless=False,
        args=args,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="pt-BR",
        timezone_id="America/Sao_Paulo",
    )
    page = context.pages[0]
    page.goto("https://web.whatsapp.com", wait_until="commit")

    print("\nAguardando carregamento da interface...")
    for _ in range(60):
        if page.is_closed():
            print("Navegador foi fechado pelo usuario.")
            break
        if page.locator("#pane-side").is_visible():
            print("\n[OK] WhatsApp Web conectado e autenticado com sucesso!")
            print("A sessao esta salva permanentemente no seu computador.")
            time.sleep(3)
            break
        if page.locator("canvas, [data-testid='qrcode']").is_visible():
            print("[INFO] QR Code detectado. Por favor, aponte a camera do seu celular para escanear.")
        time.sleep(2)

    try:
        context.close()
    except Exception:
        pass

print("\nConcluido!")
