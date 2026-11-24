#!/usr/bin/env python3
"""
GERADOR DE SESSAO WHATSAPP
--------------------------------------------------------------
Execute UMA VEZ no seu PC para autenticar o WhatsApp Web e
gerar o arquivo criptografado whatsapp_session.enc.

O formato e portavel (cookies + localStorage + IndexedDB),
para funcionar no GitHub Actions (Linux) a partir do Windows.

Passos:
  1. python 00z_gerar_sessao_zap.py
  2. Escaneie o QR Code no navegador
  3. O script gera:
       - whatsapp_session.enc  (commitar no repositorio)
       - whatsapp_key.txt      (NAO commitar - e o Secret!)
  4. Adicione a CHAVE como Secret no GitHub:
       Nome:  WHATSAPP_KEY
       Valor: conteudo do arquivo whatsapp_key.txt
     (Se a chave ja existia, nao precisa atualizar o Secret.)
  5. Commit apenas o whatsapp_session.enc:
       git add whatsapp_session.enc
       git commit -m "chore: atualiza sessao whatsapp"
       git push
--------------------------------------------------------------
"""

import os
import sys
import time
import tempfile
from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

from zap_web_sessao import (
    criptografar_payload,
    dump_sessao_pagina,
    obter_ou_criar_chave,
)

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENC_SAIDA = os.path.join(_SCRIPT_DIR, "whatsapp_session.enc")
KEY_SAIDA = os.path.join(_SCRIPT_DIR, "whatsapp_key.txt")


def _gerar_sessao(user_data_zap):
    print("=" * 60)
    print("  GERADOR DE SESSAO WHATSAPP PARA GITHUB ACTIONS")
    print("  Formato portavel (Windows -> Linux)")
    print("=" * 60)
    print(f"\nDiretorio de sessao temporario: {user_data_zap}\n")

    print("[>>] Abrindo WhatsApp Web -- escaneie o QR Code...")
    print("   Apos escanear, aguarde a tela principal carregar.")
    print("   O script fecha o navegador automaticamente.\n")

    with sync_playwright() as p:
        args_zap = [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--window-size=1920,1080",
        ]
        context = p.chromium.launch_persistent_context(
            user_data_zap,
            headless=False,
            args=args_zap,
            viewport={"width": 1920, "height": 1080},
            slow_mo=300,
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
        )
        page = context.pages[0]
        page.goto("https://web.whatsapp.com", timeout=90000)

        try:
            page.locator("canvas[aria-label*='QR'], [data-testid='qrcode']").first.wait_for(
                state="visible", timeout=30000
            )
            print("[OK] QR Code exibido.")
        except PlaywrightTimeoutError as e:
            if page.locator("#pane-side").is_visible():
                raise RuntimeError(
                    "O WhatsApp abriu conectado, sem exibir o QR Code. "
                    "O perfil do navegador nao foi iniciado como uma sessao nova."
                ) from e
            raise RuntimeError(
                "O QR Code nao apareceu em 30 segundos. Verifique a conexao e tente novamente."
            ) from e

        print("[...] Aguardando autenticacao (ate 180s)...")
        try:
            page.wait_for_selector("#pane-side", timeout=180000)
            print("\n[OK] WhatsApp autenticado com sucesso!")
            print("   Aguardando sincronizacao inicial (30s)...")
            try:
                page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
            time.sleep(30)

            print("   Verificando persistencia da sessao no localStorage...")
            tem_sessao = False
            for tentativa in range(6):
                try:
                    tem_sessao = page.evaluate("""() => {
                        try {
                            const keys = Object.keys(localStorage);
                            return keys.some(k => k.startsWith('WANoiseInfo') || (k.startsWith('W') && localStorage[k] && localStorage[k].length > 10));
                        } catch(e) { return false; }
                    }""")
                    if tem_sessao:
                        print(f"   [OK] Sessao confirmada no localStorage (tentativa {tentativa+1})")
                        break
                    print(f"   [...] Aguardando chaves de sessao... ({tentativa+1}/6)")
                    time.sleep(10)
                except Exception:
                    time.sleep(10)

            print("   Exportando cookies, localStorage e IndexedDB (formato portavel)...")
            try:
                payload = dump_sessao_pagina(page, context)
            except Exception as dump_err:
                print(f"   [AVISO] Dump completo falhou ({dump_err}). Exportando cookies/localStorage...")
                ls = {}
                try:
                    ls = page.evaluate("() => Object.fromEntries(Object.keys(localStorage).map(k => [k, localStorage.getItem(k)]))")
                except Exception:
                    pass
                payload = {
                    "format": "zap-portable-v1",
                    "storage_state": context.storage_state(),
                    "localStorage": ls or {},
                    "indexedDB": [],
                    "userAgent": page.evaluate("() => navigator.userAgent"),
                    "platform": page.evaluate("() => navigator.platform"),
                }
            n_ls = len(payload.get("localStorage") or {})
            n_idb = len(payload.get("indexedDB") or [])
            n_ck = len((payload.get("storage_state") or {}).get("cookies") or [])
            print(f"   [OK] cookies={n_ck}  localStorage={n_ls}  IndexedDB={n_idb} bancos")
            if n_ls < 1 and n_idb < 1:
                print("[ERRO] Nada foi exportado da sessao. Tente de novo.")
                context.close()
                sys.exit(1)
        except Exception as e:
            print(f"[ERRO] Timeout ou falha ao exportar sessao: {e}")
            context.close()
            sys.exit(1)

        context.close()

    chave, chave_nova = obter_ou_criar_chave(KEY_SAIDA)
    dados_enc = criptografar_payload(payload, chave)

    with open(ENC_SAIDA, "wb") as fp:
        fp.write(dados_enc)

    tamanho_enc_kb = os.path.getsize(ENC_SAIDA) / 1024
    print(f"   Arquivo criptografado: {tamanho_enc_kb:.0f} KB")

    print("\n" + "=" * 60)
    print("  [OK] CONCLUIDO!")
    print("=" * 60)
    chave_txt = chave.decode() if isinstance(chave, bytes) else chave
    if chave_nova:
        passo_secret = f"""1. Atualize o Secret no GitHub (chave NOVA gerada):
   Nome:  WHATSAPP_KEY
   Valor: {chave_txt}

   (ou copie do arquivo: {KEY_SAIDA})
"""
    else:
        passo_secret = f"""1. A chave em {KEY_SAIDA} foi reutilizada.
   Nao precisa alterar o Secret WHATSAPP_KEY no GitHub,
   a menos que o valor la esteja diferente deste arquivo.
"""
    print(f"""
PROXIMOS PASSOS:

{passo_secret}
2. Commite o arquivo criptografado:
   git add whatsapp_session.enc
   git commit -m "chore: atualiza sessao whatsapp (formato portavel)"
   git push

3. NAO commite o whatsapp_key.txt !
""")


def main():
    with tempfile.TemporaryDirectory(prefix="whatsapp_sessao_zap_") as user_data_zap:
        _gerar_sessao(user_data_zap)


if __name__ == "__main__":
    main()
