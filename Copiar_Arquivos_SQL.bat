@echo off
chcp 65001 > nul
title Sincronização BI_LEC para BI_LEC_SQL
cd /d "%~dp0"

echo ============================================================
echo   Iniciando Copia de Arquivos (BI_LEC para BI_LEC_SQL)
echo ============================================================
echo.

python copiar_arquivos_sql.py

echo.
pause

