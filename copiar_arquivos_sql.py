import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

# Garante suporte a UTF-8 no terminal Windows para evitar UnicodeEncodeError
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ==============================================================================
# MAPEAMENTO DE PASTAS (ORIGEM -> DESTINO)
# ==============================================================================
MAPEAMENTO = [
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\01_ELF_Diario",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\01_ELF_Diario",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\02_PDL_Analitico",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\02_PDL_Analitico",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\03_N_Lib_Fat_Diario",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\03_N_Lib_Fat_Diario",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\04_N_Visitado_Diario",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\04_N_Visitado_Diario",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\05_N_Visitado_Historico",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\05_N_Visitado_Historico",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\06_Impedimentos",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\06_Impedimento",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\07_Entregas",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\07_Entrega",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\08_Leituras_Correcao",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\08_Leituras_Correcao",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\10_Aviso_Desliga",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\10_Aviso_Desliga",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\11_Rastro_Entrega",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\11_Rastro_Entrega",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\12_Validacao_Malote",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\12_Validacao_Malote",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\15_Planeja_Bases",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\15_Planieja_Bases",
    ),
    (
        r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\16_Notas_Servico",
        r"C:\Users\paulo.janio\OneDrive - ENGELMIG ENERGIA LTDA\BI_LEC_SQL\16_Notas_Servico",
    ),
]


def copiar_pastas(sobrescrever: bool = False) -> None:
    inicio = time.time()
    data_inicio = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    print("=" * 80)
    print("INICIANDO SINCRONIZACAO DE ARQUIVOS (BI_LEC -> BI_LEC_SQL)")
    print(f"Data/Hora de Inicio: {data_inicio}")
    print(f"Modo: {'SOBRESCREVER TUDO' if sobrescrever else 'IGNORAR ARQUIVOS JA EXISTENTES'}")
    print("=" * 80)

    total_copiados_geral = 0
    total_ignorados_geral = 0
    total_erros_geral = 0

    for idx, (origem_str, destino_str) in enumerate(MAPEAMENTO, start=1):
        origem = Path(origem_str)
        destino = Path(destino_str)

        nome_origem = origem.name
        nome_destino = destino.name
        label = f"[{idx:02d}/{len(MAPEAMENTO):02d}] {nome_origem} -> {nome_destino}"

        print(f"\nPasta: {label}")

        if not origem.exists():
            print(f"   [AVISO] Pasta de origem nao encontrada: {origem}")
            continue

        destino.mkdir(parents=True, exist_ok=True)

        itens = [item for item in origem.iterdir() if item.is_file()]

        if not itens:
            print("   [INFO] Nenhum arquivo encontrado na pasta de origem.")
            continue

        copiados_pasta = 0
        ignorados_pasta = 0
        erros_pasta = 0

        for item in itens:
            destino_arquivo = destino / item.name

            if destino_arquivo.exists() and not sobrescrever:
                ignorados_pasta += 1
                continue

            try:
                # copy2 preserva datas de criacao e modificacao do arquivo
                shutil.copy2(item, destino_arquivo)
                copiados_pasta += 1
            except Exception as e:
                erros_pasta += 1
                print(f"   [ERRO] Ao copiar '{item.name}': {e}")

        total_copiados_geral += copiados_pasta
        total_ignorados_geral += ignorados_pasta
        total_erros_geral += erros_pasta

        print(
            f"   -> Copiado(s): {copiados_pasta} | "
            f"Ignorado(s) (ja existiam): {ignorados_pasta} | "
            f"Erro(s): {erros_pasta}"
        )

    tempo_total = time.time() - inicio
    print("\n" + "=" * 80)
    print("SINCRONIZACAO CONCLUIDA")
    print(f"Tempo total: {tempo_total:.2f} segundos")
    print(f"Total de arquivos copiados:           {total_copiados_geral}")
    print(f"Total de arquivos ignorados (exist.): {total_ignorados_geral}")
    if total_erros_geral > 0:
        print(f"Total de erros:                       {total_erros_geral}")
    print("=" * 80)


if __name__ == "__main__":
    forcar = "--force" in sys.argv or "-f" in sys.argv
    copiar_pastas(sobrescrever=forcar)

