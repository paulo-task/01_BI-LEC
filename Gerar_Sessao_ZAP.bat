@echo off
title Autenticar e Exportar Sessao WhatsApp para GitHub
cd /d "%~dp0"
echo ======================================================
echo  AUTENTICAR SESSAO WHATSAPP PARA O GITHUB ACTIONS
echo ======================================================
python 00z_gerar_sessao.py
pause
