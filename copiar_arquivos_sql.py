import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

# Garante suporte a UTF-8 no terminal Windows para evitar UnicodeEncodeError
if sys.platform == "win32":
    try:
        if sys.stdout is not None:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if sys.stderr is not None:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).resolve().parent
LOGS_DIR = SCRIPT_DIR / "Logs"
LOG_FILE = LOGS_DIR / "copiar_arquivos_sql.log"

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


class DualLogger:
    """Grava as saídas no arquivo de log e também no console (se houver)."""
    def __init__(self, log_path: Path):
        self.terminal = sys.stdout
        self.log_path = log_path
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        self.log_file = open(self.log_path, "a", encoding="utf-8", errors="replace")

    def write(self, message):
        if self.terminal:
            try:
                self.terminal.write(message)
            except Exception:
                pass
        self.log_file.write(message)
        self.log_file.flush()

    def flush(self):
        if self.terminal:
            try:
                self.terminal.flush()
            except Exception:
                pass
        self.log_file.flush()

    def close(self):
        try:
            self.log_file.close()
        except Exception:
            pass


def copiar_pastas(sobrescrever_tudo: bool = False) -> None:
    logger = DualLogger(LOG_FILE)
    sys.stdout = logger

    try:
        inicio = time.time()
        agora = datetime.now()
        hoje = agora.date()
        hoje_str = agora.strftime("%Y_%m_%d")
        data_inicio = agora.strftime("%d/%m/%Y %H:%M:%S")

        print("=" * 80)
        print("INICIANDO SINCRONIZACAO DE ARQUIVOS (BI_LEC -> BI_LEC_SQL)")
        print(f"Data/Hora de Inicio: {data_inicio}")
        print("Regra: Ignorar dias anteriores ja existentes | Atualizar arquivos do dia")
        print("=" * 80)

        total_copiados_novos = 0
        total_atualizados_hoje = 0
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

            novos_pasta = 0
            atualizados_pasta = 0
            ignorados_pasta = 0
            erros_pasta = 0

            for item in itens:
                destino_arquivo = destino / item.name

                stat_origem = item.stat()
                mtime_origem = stat_origem.st_mtime
                dt_origem = datetime.fromtimestamp(mtime_origem)

                # Verifica se é um arquivo do dia atual (por data de modificação ou nome do arquivo)
                eh_dia_atual = (hoje_str in item.name) or (dt_origem.date() == hoje)

                # Caso 1: Arquivo não existe no destino -> Sempre copia
                if not destino_arquivo.exists():
                    try:
                        shutil.copy2(item, destino_arquivo)
                        novos_pasta += 1
                    except Exception as e:
                        erros_pasta += 1
                        print(f"   [ERRO] Ao copiar novo '{item.name}': {e}")
                    continue

                # Caso 2: Arquivo já existe no destino
                if sobrescrever_tudo:
                    deve_copiar = True
                    motivo = "forçado"
                elif eh_dia_atual:
                    # No dia atual: verifica se a origem é mais recente que o destino
                    stat_destino = destino_arquivo.stat()
                    mtime_destino = stat_destino.st_mtime
                    size_origem = stat_origem.st_size
                    size_destino = stat_destino.st_size

                    # Copia se a origem foi modificada depois do destino ou se o tamanho mudou
                    if (mtime_origem > mtime_destino + 1) or (size_origem != size_destino):
                        deve_copiar = True
                        motivo = f"atualizado hoje (Origem: {dt_origem.strftime('%H:%M:%S')} vs Destino: {datetime.fromtimestamp(mtime_destino).strftime('%H:%M:%S')})"
                    else:
                        deve_copiar = False
                else:
                    # Dias anteriores já existentes: ignora
                    deve_copiar = False

                if deve_copiar:
                    try:
                        shutil.copy2(item, destino_arquivo)
                        atualizados_pasta += 1
                        print(f"   -> [SUBSTITUIDO] {item.name} ({motivo})")
                    except Exception as e:
                        erros_pasta += 1
                        print(f"   [ERRO] Ao atualizar '{item.name}': {e}")
                else:
                    ignorados_pasta += 1

            total_copiados_novos += novos_pasta
            total_atualizados_hoje += atualizados_pasta
            total_ignorados_geral += ignorados_pasta
            total_erros_geral += erros_pasta

            print(
                f"   Resumo: {novos_pasta} novo(s) | "
                f"{atualizados_pasta} atualizado(s) do dia | "
                f"{ignorados_pasta} ignorado(s) | "
                f"{erros_pasta} erro(s)"
            )

        tempo_total = time.time() - inicio
        print("\n" + "=" * 80)
        print("SINCRONIZACAO CONCLUIDA")
        print(f"Tempo total: {tempo_total:.2f} segundos")
        print(f"Novos arquivos copiados:              {total_copiados_novos}")
        print(f"Arquivos do dia atualizados:          {total_atualizados_hoje}")
        print(f"Arquivos ignorados (dias ant./iguais):{total_ignorados_geral}")
        if total_erros_geral > 0:
            print(f"Total de erros:                       {total_erros_geral}")
        print("=" * 80 + "\n")
    finally:
        sys.stdout = logger.terminal
        logger.close()


if __name__ == "__main__":
    forcar = "--force" in sys.argv or "-f" in sys.argv
    copiar_pastas(sobrescrever_tudo=forcar)
