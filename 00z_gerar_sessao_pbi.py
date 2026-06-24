#!/usr/bin/env python3
"""
GERADOR DE SESSAO POWER BI PARA GITHUB
======================================
Exporta cookies via Playwright storage_state (portável Windows -> Linux).

Uso:  python 00z_gerar_sessao_pbi.py

Gera:
  powerbi_session.enc  -> commitar
  powerbi_key.txt      -> secret POWERBI_KEY no GitHub (NAO commitar)
"""

import json
import os
import sys
import time
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright
from cryptography.fernet import Fernet

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_PBI = os.path.join(_SCRIPT_DIR, "dados_pbi")
PERFIL_00C = r"C:\Temp\dados_navegador_print"
ENC_SAIDA = os.path.join(_SCRIPT_DIR, "powerbi_session.enc")
KEY_SAIDA = os.path.join(_SCRIPT_DIR, "powerbi_key.txt")
STATE_TMP = os.path.join(_SCRIPT_DIR, "powerbi_state_tmp.json")

URL_POWERBI = (
    "https://app.powerbi.com/groups/33331c64-94a0-477c-b682-9f40a7ac809b/reports/"
    "ff3f2d1f-9433-4060-b072-07b666de8da0/9592b20a8d6c05c3d407?experience=power-bi"
)


def perfil_tem_dados(pasta):
    if not os.path.isdir(pasta):
        return False
    for rel in ("Default/Cookies", "Default/Network/Cookies", "Default/Preferences", "Local State"):
        if os.path.exists(os.path.join(pasta, *rel.split("/"))):
            return True
    return False


def preparar_perfil():
    Path(USER_DATA_PBI).mkdir(parents=True, exist_ok=True)
    if perfil_tem_dados(USER_DATA_PBI):
        return "dados_pbi"
    if perfil_tem_dados(PERFIL_00C):
        print(f"Copiando perfil de {PERFIL_00C} ...")
        if os.path.exists(USER_DATA_PBI):
            shutil.rmtree(USER_DATA_PBI)
        shutil.copytree(PERFIL_00C, USER_DATA_PBI, ignore=shutil.ignore_patterns("Cache", "Code Cache", "GPUCache"))
        return "C:\\Temp\\dados_navegador_print"
    return "dados_pbi (novo)"


def exportar_storage_state():
    origem = preparar_perfil()
    print(f"Perfil: {origem}")
    print("\n[>>] Abrindo Power BI — aguarde a aba 'ELF Hora' (faca login se pedir)...\n")

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
        try:
            page.get_by_role("tab", name="ELF Hora").wait_for(state="visible", timeout=180000)
            print("[OK] Relatorio carregou!")
            time.sleep(5)
        except Exception:
            print("[ERRO] Relatorio nao carregou. Faca login e rode de novo.")
            context.close()
            sys.exit(1)

        context.storage_state(path=STATE_TMP)
        context.close()

    with open(STATE_TMP, encoding="utf-8") as fp:
        state = json.load(fp)
    n_cookies = len(state.get("cookies", []))
    if n_cookies < 3:
        print(f"[ERRO] Poucos cookies exportados ({n_cookies}). Login incompleto.")
        sys.exit(1)

    print(f"[OK] {n_cookies} cookies exportados (formato portavel para Linux/GitHub)")

    chave = Fernet.generate_key()
    with open(STATE_TMP, "rb") as fp:
        dados_enc = Fernet(chave).encrypt(fp.read())
    os.remove(STATE_TMP)

    with open(ENC_SAIDA, "wb") as fp:
        fp.write(dados_enc)
    with open(KEY_SAIDA, "w") as fp:
        fp.write(chave.decode())

    print(f"   powerbi_session.enc: {os.path.getsize(ENC_SAIDA) / 1024:.0f} KB")
    print(f"   powerbi_key.txt gerado")


def main():
    print("=" * 60)
    print("  EXPORTAR SESSAO POWER BI  ->  GITHUB (formato portatil)")
    print("=" * 60 + "\n")
    exportar_storage_state()
    print("\n" + "=" * 60)
    print("  CONCLUIDO!")
    print("=" * 60)
    print("""
1. GitHub Secret POWERBI_KEY = conteudo de powerbi_key.txt
   (se a chave mudou, atualize o secret)

2. git add powerbi_session.enc && git commit && git push

3. Dispare o workflow no GitHub
""")


if __name__ == "__main__":
    main()
