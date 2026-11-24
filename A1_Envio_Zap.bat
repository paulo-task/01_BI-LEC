@echo off
title Automacao Power BI + WhatsApp (BI-LEC)
cd /d "%~dp0"
echo ======================================================
echo  EXECUTANDO RELATORIO POWER BI + WHATSAPP
echo ======================================================
C:\Python314\python.exe A1_Envio_Zap.py
set RET_CODE=%ERRORLEVEL%
if %RET_CODE% NEQ 0 (
    echo.
    echo ======================================================
    echo [ERRO] Ocorreu uma falha na execucao. Codigo: %RET_CODE%
    echo ======================================================
    pause
    exit /b %RET_CODE%
)
echo.
echo ======================================================
echo [OK] Relatorio capturado e enviado com sucesso!
echo ======================================================
timeout /t 5 >nul
exit /b 0
