#!/usr/bin/env python3
"""
GERADOR DE SESSÃO WHATSAPP
─────────────────────────────────────────────────────────────────
Execute este script UMA VEZ no seu PC para autenticar o WhatsApp Web
e gerar o Secret WHATSAPP_SESSION para o GitHub Actions.

Passos:
  1. python 00z_gerar_sessao.py
  2. Escaneie o QR Code que aparecer no navegador
  3. Aguarde o script finalizar e gerar o arquivo sessao_base64.txt
  4. Copie o conteúdo do arquivo e cole como Secret no GitHub:
     Repositório → Settings → Secrets → Actions → New secret
     Nome: WHATSAPP_SESSION
─────────────────────────────────────────────────────────────────
"""

import os
import sys
import time
import base64
import zipfile
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright

# Pasta dados_zap dentro do próprio repositório (funciona em qualquer PC)
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_ZAP = os.path.join(_SCRIPT_DIR, "dados_zap")
ZIP_SAIDA     = os.path.join(_SCRIPT_DIR, "dados_zap_sessao.zip")
TXT_SAIDA     = os.path.join(_SCRIPT_DIR, "sessao_base64.txt")


def main():
    print("=" * 60)
    print("  GERADOR DE SESSÃO WHATSAPP PARA GITHUB ACTIONS")
    print("=" * 60)
    print(f"\nDiretório de sessão: {USER_DATA_ZAP}\n")

    Path(USER_DATA_ZAP).mkdir(parents=True, exist_ok=True)

    print("[>>] Abrindo WhatsApp Web -- escaneie o QR Code...")
    print("   Apos escanear, aguarde a tela principal carregar.")
    print("   O script fecha o navegador automaticamente.\n")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_ZAP,
            headless=False,
            args=["--start-maximized"],
            no_viewport=True,
            slow_mo=500,
        )
        page = context.pages[0]
        page.goto("https://web.whatsapp.com", timeout=60000)

        print("[...] Aguardando autenticacao (ate 120s)...")
        try:
            page.wait_for_selector("#pane-side", timeout=120000)
            print("[OK] WhatsApp autenticado com sucesso!")
            print("   Aguardando 5s para estabilizar a sessao...")
            time.sleep(5)
        except Exception:
            print("[ERRO] Timeout -- QR Code nao foi escaneado a tempo.")
            context.close()
            sys.exit(1)

        context.close()

    # Compacta a pasta de sessão
    print(f"\n[ZIP] Compactando sessao em {ZIP_SAIDA}...")
    if os.path.exists(ZIP_SAIDA):
        os.remove(ZIP_SAIDA)

    with zipfile.ZipFile(ZIP_SAIDA, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(USER_DATA_ZAP):
            # Ignora arquivos de lock que o Chromium cria
            dirs[:] = [d for d in dirs if d not in ["SingletonLock", "SingletonSocket"]]
            for file in files:
                if file in ("SingletonLock", "SingletonSocket", "lockfile"):
                    continue
                caminho_abs = os.path.join(root, file)
                caminho_rel = os.path.relpath(caminho_abs, USER_DATA_ZAP)
                zf.write(caminho_abs, caminho_rel)

    tamanho_mb = os.path.getsize(ZIP_SAIDA) / (1024 * 1024)
    print(f"   Tamanho: {tamanho_mb:.1f} MB")

    # Codifica em base64
    print(f"[B64] Codificando em base64 -> {TXT_SAIDA}...")
    with open(ZIP_SAIDA, "rb") as f:
        conteudo_b64 = base64.b64encode(f.read()).decode("utf-8")

    with open(TXT_SAIDA, "w") as f:
        f.write(conteudo_b64)

    tamanho_kb = os.path.getsize(TXT_SAIDA) / 1024
    print(f"   Tamanho base64: {tamanho_kb:.0f} KB")

    print("\n" + "=" * 60)
    print("  [OK] CONCLUIDO!")
    print("=" * 60)
    print(f"""
PROXIMO PASSO -- Adicione o Secret no GitHub:

  1. Abra:  github.com -> seu repositorio
  2. Va em: Settings -> Secrets and variables -> Actions
  3. Clique: New repository secret
  4. Nome:   WHATSAPP_SESSION
  5. Valor:  cole o conteudo do arquivo '{TXT_SAIDA}'

ATENCAO: O arquivo gerado pode ter varios KB -- use Ctrl+A para
    selecionar tudo no bloco de notas antes de copiar.

ATENCAO: Repita este processo se o WhatsApp deslogar (troca de
    celular, inatividade longa, etc).
""")


if __name__ == "__main__":
    main()
