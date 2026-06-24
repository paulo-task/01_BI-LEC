#!/usr/bin/env python3
"""
GERADOR DE SESSAO POWER BI PARA GITHUB
======================================
Exporta o login do Chrome do seu PC (cookies) para o GitHub Actions.

Uso:
  python 00z_gerar_sessao_pbi.py

Gera:
  powerbi_session.enc  -> commitar no repositorio
  powerbi_key.txt      -> NAO commitar; colar no GitHub Secret POWERBI_KEY

Se voce ja roda prints pelo 00c_Print_Telas.py, o script usa automaticamente
o perfil em C:\\Temp\\dados_navegador_print.
"""

import os
import sys
import time
import zipfile
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright
from cryptography.fernet import Fernet

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_PBI = os.path.join(_SCRIPT_DIR, "dados_pbi")
PERFIL_00C = r"C:\Temp\dados_navegador_print"
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


def perfil_tem_dados(pasta):
    if not os.path.isdir(pasta):
        return False
    indicadores = [
        os.path.join(pasta, "Default", "Cookies"),
        os.path.join(pasta, "Default", "Network", "Cookies"),
        os.path.join(pasta, "Default", "Preferences"),
        os.path.join(pasta, "Local State"),
    ]
    return any(os.path.exists(p) for p in indicadores)


def escolher_pasta_perfil():
    if perfil_tem_dados(USER_DATA_PBI):
        return USER_DATA_PBI, "dados_pbi (00z_capturar_envia)"
    if perfil_tem_dados(PERFIL_00C):
        return PERFIL_00C, "C:\\Temp\\dados_navegador_print (00c_Print_Telas)"
    return USER_DATA_PBI, "dados_pbi (novo — vai abrir o navegador)"


def compactar_sessao(pasta_origem, zip_destino):
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
                    contagem += 1
                except (PermissionError, OSError):
                    pass
    return contagem


def exportar_enc(pasta_origem):
    if os.path.exists(ZIP_TMP):
        os.remove(ZIP_TMP)

    qtd = compactar_sessao(pasta_origem, ZIP_TMP)
    if qtd == 0:
        print("[ERRO] Nenhum arquivo exportado — perfil vazio ou bloqueado.")
        sys.exit(1)

    tamanho_mb = os.path.getsize(ZIP_TMP) / (1024 * 1024)
    print(f"   {qtd} arquivos | {tamanho_mb:.1f} MB compactado")

    chave = Fernet.generate_key()
    with open(ZIP_TMP, "rb") as fp:
        dados_enc = Fernet(chave).encrypt(fp.read())
    os.remove(ZIP_TMP)

    with open(ENC_SAIDA, "wb") as fp:
        fp.write(dados_enc)

    with open(KEY_SAIDA, "w") as fp:
        fp.write(chave.decode())

    print(f"   powerbi_session.enc: {os.path.getsize(ENC_SAIDA) / 1024:.0f} KB")
    print(f"   powerbi_key.txt gerado (NAO commitar)")


def validar_no_navegador(pasta_destino):
    """Abre Power BI com o perfil escolhido e confirma que ELF Hora carrega."""
    print("\n[>>] Validando sessao no navegador...")
    print("     Se pedir login, faca manualmente e aguarde a aba ELF Hora.\n")

    destino = USER_DATA_PBI
    if pasta_destino != destino:
        if os.path.exists(destino):
            shutil.rmtree(destino)
        shutil.copytree(pasta_destino, destino, ignore=shutil.ignore_patterns(*EXCLUIR_DIRS))

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            destino,
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
            print("[OK] Relatorio carregou — sessao valida!")
            time.sleep(5)
        except Exception:
            print("[ERRO] Relatorio nao carregou. Faca login e rode de novo.")
            context.close()
            sys.exit(1)
        context.close()


def main():
    print("=" * 60)
    print("  EXPORTAR SESSAO POWER BI  ->  GITHUB ACTIONS")
    print("=" * 60)

    pasta, origem = escolher_pasta_perfil()
    print(f"\nPerfil detectado: {origem}")

    if perfil_tem_dados(pasta):
        print("[OK] Cookies encontrados — exportando sessao do PC...")
        exportar_enc(pasta)
    else:
        print("[...] Perfil vazio — abrindo Power BI para voce logar...")
        Path(USER_DATA_PBI).mkdir(parents=True, exist_ok=True)
        validar_no_navegador(USER_DATA_PBI)
        exportar_enc(USER_DATA_PBI)

    print("\n" + "=" * 60)
    print("  CONCLUIDO!")
    print("=" * 60)

    with open(KEY_SAIDA) as fp:
        chave = fp.read().strip()

    print(f"""
PROXIMOS PASSOS (uma vez):

1. GitHub -> Settings -> Secrets -> Actions -> New secret
   Nome:  POWERBI_KEY
   Valor: (conteudo de powerbi_key.txt)

2. Commit do arquivo criptografado:
   git add powerbi_session.enc
   git commit -m "chore: adiciona sessao power bi para github"
   git push

3. Dispare o workflow no GitHub — vai usar o mesmo login do PC.

Quando a sessao expirar (semanas/meses), rode este script de novo.
NAO commite powerbi_key.txt
""")


if __name__ == "__main__":
    main()
