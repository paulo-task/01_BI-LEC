import hashlib
import os
import tempfile

import pandas as pd


COLUNAS_INTEIRAS = (
    "TAREFA",
    "T. VISITADA",
    "T. DISTRIB",
    "QTD FAT",
    "QTD FAT DISTRIB",
    "QTD FAT ENTREGUES",
    "QTD FAT DEVOLVIDAS",
)


def caminhos_arquivos_quinzenais(
    diretorio_destino: str, nome_arquivo: str
) -> tuple[str, str]:
    projeto = os.path.normcase(os.path.abspath(os.path.dirname(__file__)))
    destino = os.path.normcase(os.path.abspath(diretorio_destino))
    identificador = hashlib.sha256(
        f"{projeto}|{destino}".encode("utf-8")
    ).hexdigest()[:16]
    nome_base = os.path.splitext(os.path.basename(nome_arquivo))[0]
    diretorio_temporario = os.path.join(
        tempfile.gettempdir(), "planeja_bases", identificador, nome_base
    )
    return (
        os.path.join(diretorio_temporario, f"{nome_base}_Q1.csv"),
        os.path.join(diretorio_temporario, f"{nome_base}_Q2.csv"),
    )


def remover_arquivos_quinzenais(caminho_q1: str, caminho_q2: str) -> None:
    for caminho in (caminho_q1, caminho_q2):
        if os.path.exists(caminho):
            os.remove(caminho)

    diretorio_mes = os.path.dirname(caminho_q1)
    diretorio_projeto = os.path.dirname(diretorio_mes)
    diretorio_temporario = os.path.dirname(diretorio_projeto)
    for diretorio in (diretorio_mes, diretorio_projeto, diretorio_temporario):
        if os.path.isdir(diretorio) and not os.listdir(diretorio):
            os.rmdir(diretorio)


def normalizar_colunas_inteiras(df: pd.DataFrame) -> pd.DataFrame:
    resultado = df.copy()

    for coluna in COLUNAS_INTEIRAS:
        valores = resultado[coluna].replace(r"^\s*$", pd.NA, regex=True)
        numeros = pd.to_numeric(valores, errors="coerce")

        invalidos = valores.notna() & numeros.isna()
        if invalidos.any():
            raise ValueError(f"A coluna '{coluna}' contém valores não numéricos.")

        fracionarios = numeros.notna() & numeros.mod(1).ne(0)
        if fracionarios.any():
            raise ValueError(f"A coluna '{coluna}' contém valores fracionários.")

        resultado[coluna] = numeros.astype("Int64")

    return resultado
