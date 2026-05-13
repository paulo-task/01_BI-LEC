#!/usr/bin/env python3
"""
GERADOR DE SESSAO WHATSAPP
--------------------------------------------------------------
Execute UMA VEZ no seu PC para autenticar o WhatsApp Web e
gerar o arquivo criptografado whatsapp_session.enc.

Passos:
  1. python 00z_gerar_sessao.py
  2. Escaneie o QR Code no navegador
  3. O script gera:
       - whatsapp_session.enc  (commitar no repositorio)
       - whatsapp_key.txt      (NAO commitar - e o Secret!)
  4. Adicione a CHAVE como Secret no GitHub:
       Nome:  WHATSAPP_KEY
       Valor: conteudo do arquivo whatsapp_key.txt
  5. Commit apenas o whatsapp_session.enc:
       git add whatsapp_session.enc
       git commit -m "chore: atualiza sessao whatsapp"
       git push
--------------------------------------------------------------
"""

import os
import sys
import time
import zipfile
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright
from cryptography.fernet import Fernet

# Pasta dados_zap dentro do repositorio
_SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
USER_DATA_ZAP = os.path.join(_SCRIPT_DIR, "dados_zap")
ENC_SAIDA     = os.path.join(_SCRIPT_DIR, "whatsapp_session.enc")
KEY_SAIDA     = os.path.join(_SCRIPT_DIR, "whatsapp_key.txt")
ZIP_TMP       = os.path.join(_SCRIPT_DIR, "dados_zap_tmp.zip")

# Pastas do Chromium que NAO precisam ser salvas (cache, etc.)
EXCLUIR_DIRS = {
    "Cache", "Code Cache", "GPUCache", "ShaderCache",
    "DawnGraphiteCache", "DawnWebGPUCache", "GrShaderCache",
    "blob_storage", "CrashpadMetrics-active.pma",
    "component_crx_cache", "hyphen-data",
    "Media Cache", "Service Worker", "Session Storage",
    "VideoDecodeStats", "File System", "WebStorage",
}
EXCLUIR_ARQUIVOS = {
    "SingletonLock", "SingletonSocket", "lockfile",
    "SingletonCookie", "LOCK", "LOG", "LOG.old",
}


def compactar_sessao(pasta_origem, zip_destino):
    """Compacta apenas os arquivos essenciais da sessao."""
    total_bytes = 0
    contagem = 0
    with zipfile.ZipFile(zip_destino, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(pasta_origem):
            # Remove dirs de cache da lista de visita
            dirs[:] = [d for d in dirs if d not in EXCLUIR_DIRS]
            for file in files:
                if file in EXCLUIR_ARQUIVOS:
                    continue
                caminho_abs = os.path.join(root, file)
                caminho_rel = os.path.relpath(caminho_abs, pasta_origem)
                try:
                    zf.write(caminho_abs, caminho_rel)
                    total_bytes += os.path.getsize(caminho_abs)
                    contagem += 1
                except (PermissionError, OSError):
                    pass  # ignora arquivos bloqueados
    return contagem, total_bytes


def main():
    print("=" * 60)
    print("  GERADOR DE SESSAO WHATSAPP PARA GITHUB ACTIONS")
    print("=" * 60)
    print(f"\nDiretorio de sessao: {USER_DATA_ZAP}\n")

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

    # Compacta apenas arquivos essenciais
    print(f"\n[ZIP] Compactando sessao (sem cache)...")
    if os.path.exists(ZIP_TMP):
        os.remove(ZIP_TMP)

    qtd, total = compactar_sessao(USER_DATA_ZAP, ZIP_TMP)
    tamanho_mb = os.path.getsize(ZIP_TMP) / (1024 * 1024)
    print(f"   {qtd} arquivos | {tamanho_mb:.1f} MB compactado")

    # Criptografa com Fernet
    print("[ENC] Criptografando...")
    chave = Fernet.generate_key()
    f = Fernet(chave)

    with open(ZIP_TMP, "rb") as fp:
        dados_zip = fp.read()

    dados_enc = f.encrypt(dados_zip)

    with open(ENC_SAIDA, "wb") as fp:
        fp.write(dados_enc)

    tamanho_enc_kb = os.path.getsize(ENC_SAIDA) / 1024
    print(f"   Arquivo criptografado: {tamanho_enc_kb:.0f} KB")

    # Salva a chave
    with open(KEY_SAIDA, "w") as fp:
        fp.write(chave.decode())

    # Remove zip temporario
    os.remove(ZIP_TMP)

    print("\n" + "=" * 60)
    print("  [OK] CONCLUIDO!")
    print("=" * 60)
    print(f"""
PROXIMOS PASSOS:

1. Adicione a CHAVE como Secret no GitHub:
   URL: github.com -> seu repositorio -> Settings
        -> Secrets and variables -> Actions
        -> New repository secret
   Nome:  WHATSAPP_KEY
   Valor: {chave.decode()}

   (ou copie do arquivo: {KEY_SAIDA})

2. Commite o arquivo criptografado (pode commitar com seguranca):
   git add whatsapp_session.enc
   git commit -m "chore: atualiza sessao whatsapp"
   git push

3. NAO commite o whatsapp_key.txt !
""")


if __name__ == "__main__":
    main()
