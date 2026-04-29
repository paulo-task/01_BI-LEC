@echo off
title Executando Rotina de Prints e WhatsApp - Engelmig
echo ======================================================
echo INICIANDO PROCESSO DE CAPTURA E ENVIO
echo ======================================================

:: 1. Entra na pasta onde o seu script Python está salvo
cd /d "C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\Python"

:: 2. Executa o script
:: Se o seu comando for apenas 'python', mantenha assim. 
:: Se usar 'python3', altere o nome abaixo.
python 00c_Print_Telas.py

echo ======================================================
echo PROCESSO FINALIZADO!
echo ======================================================
exit