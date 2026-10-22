@echo off
chcp 65001 > nul
cd /d "%~dp0"

if "%~1"=="--silent" goto :run_silent
if "%~1"=="-s" goto :run_silent

title Sincronizacao BI_LEC para BI_LEC_SQL
echo ============================================================
echo   Iniciando Copia de Arquivos (BI_LEC para BI_LEC_SQL)
echo ============================================================
echo.

python -u copiar_arquivos_sql.py

echo.
pause
goto :eof

:run_silent
python -u copiar_arquivos_sql.py
exit /b %ERRORLEVEL%
