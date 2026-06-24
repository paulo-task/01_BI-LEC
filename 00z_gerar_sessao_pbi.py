#!/usr/bin/env python3
"""
GERADOR DE SESSAO POWER BI
--------------------------------------------------------------
Execute no PC DEPOIS de logar no Power BI pelo menos uma vez
(rodando 00z_capturar_envia.py localmente ou abrindo o relatório).

Passos:
  1. python 00z_gerar_sessao_pbi.py
  2. Aguarde o relatório carregar no navegador
  3. O script gera:
       - powerbi_session.enc  (commitar no repositorio)
       - powerbi_key.txt      (NAO commitar - e o Secret!)
  4. Adicione a CHAVE como Secret no GitHub:
       Nome:  POWERBI_KEY
       Valor: conteudo do arquivo powerbi_key.txt
  5. Commit apenas o powerbi_session.enc:
       git add powerbi_session.enc
       git commit -m "chore: atualiza sessao power bi"
       git push
--------------------------------------------------------------
"""

import os
import sys
import time
import zipfile
from pathlib import Path
from playwright.sync_api import sync_playwright
from cryptography.fernet import Fernet

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_PBI = os.path.join(_SCRIPT_DIR, "dados_pbi")
ENC_SAIDA = os.path.join(_SCRIPT_DIR, "powerbi_session.enc")
KEY_SAIDA = os.path.join(_SCRIPT_DIR, "powerbi_key.txt")
ZIP_TMP = os.path.join(_SCRIPT_DIR, "dados_pbi_tmp.zip")

URL_POWERBI = (
    "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/"
    "ff3f2d1f-9433-4060-b072-07b666de8da0/9592b20a8d6c05c3d407?experience=power-bi"
)

EXCLUIR_DIRS = {
    "Cache", "Code Cache", "GPUCache", "ShaderCache",
    "DawnGraphiteCache", "DawnWebGPUCache", "GrShaderCache",
    "blob_storage", "CrashpadMetrics-active.pma",
    "component_crx_cache", "hyphen-data",
}
EXCLUIR_ARQUIVOS = {
    "SingletonLock", "SingletonSocket", "lockfile",
    "SingletonCookie", "LOCK", "LOG", "LOG.old",
}


def compactar_sessao(pasta_origem, zip_destino):
    total_bytes = 0
    contagem = 0
    with zipfile.ZipFile(zip_destino, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(pasta_origem):
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
                    pass
    return contagem, total_bytes


def main():
    print("=" * 60)
    print("  GERADOR DE SESSAO POWER BI PARA GITHUB ACTIONS")
    print("=" * 60)
    print(f"\nDiretorio de sessao: {USER_DATA_PBI}\n")

    Path(USER_DATA_PBI).mkdir(parents=True, exist_ok=True)

    print("[>>] Abrindo Power BI — confirme que o relatorio carrega sem pedir login...")
    print("   Se pedir login, faca manualmente e aguarde a aba 'ELF Hora' aparecer.\n")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_PBI,
            headless=False,
            args=["--start-maximized", "--window-size=1920,1080"],
            viewport={"width": 1920, "height": 1080},
            slow_mo=300,
            locale="pt-BR",
        )
        page = context.pages[0]
        page.goto(URL_POWERBI, timeout=120000)

        print("[...] Aguardando relatorio (ate 180s)...")
        try:
            page.get_by_role("tab", name="ELF Hora").wait_for(state="visible", timeout=180000)
            print("[OK] Power BI autenticado — aba ELF Hora visivel!")
            print("   Aguardando 5s para estabilizar a sessao...")
            time.sleep(5)
        except Exception:
            print("[ERRO] Relatorio nao carregou — faca login manualmente e rode de novo.")
            context.close()
            sys.exit(1)

        context.close()

    print(f"\n[ZIP] Compactando sessao (sem cache)...")
    if os.path.exists(ZIP_TMP):
        os.remove(ZIP_TMP)

    qtd, _ = compactar_sessao(USER_DATA_PBI, ZIP_TMP)
    tamanho_mb = os.path.getsize(ZIP_TMP) / (1024 * 1024)
    print(f"   {qtd} arquivos | {tamanho_mb:.1f} MB compactado")

    print("[ENC] Criptografando...")
    chave = Fernet.generate_key()
    f = Fernet(chave)

    with open(ZIP_TMP, "rb") as fp:
        dados_enc = f.encrypt(fp.read())

    with open(ENC_SAIDA, "wb") as fp:
        fp.write(dados_enc)

    tamanho_enc_kb = os.path.getsize(ENC_SAIDA) / 1024
    print(f"   Arquivo criptografado: {tamanho_enc_kb:.0f} KB")

    with open(KEY_SAIDA, "w") as fp:
        fp.write(chave.decode())

    os.remove(ZIP_TMP)

    print("\n" + "=" * 60)
    print("  [OK] CONCLUIDO!")
    print("=" * 60)
    print(f"""
PROXIMOS PASSOS:

1. Adicione a CHAVE como Secret no GitHub:
   Nome:  POWERBI_KEY
   Valor: {chave.decode()}

   (ou copie do arquivo: {KEY_SAIDA})

2. Commite o arquivo criptografado:
   git add powerbi_session.enc
   git commit -m "chore: atualiza sessao power bi"
   git push

3. NAO commite o powerbi_key.txt !
""")


if __name__ == "__main__":
    main()
