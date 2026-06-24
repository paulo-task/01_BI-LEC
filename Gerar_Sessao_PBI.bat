@echo off
title Exportar sessao Power BI para GitHub
cd /d "%~dp0"
echo ======================================================
echo  EXPORTAR LOGIN DO PC PARA O GITHUB
echo ======================================================
python 00z_gerar_sessao_pbi.py
pause
