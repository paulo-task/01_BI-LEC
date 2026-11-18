@echo off
:: O comando abaixo faz o script 'pular' para a pasta onde ele está salvo
cd /d "%~dp0"

cls
echo ======================================================
echo   SOLICITACAO FINAL DE SEMANA (PDL E NAO VISITADAS)
echo ======================================================

echo [PASSO 1] Solicitando 02_PDL_Analitico...
:: Removi o 'start' para o log aparecer aqui
python 02_PDL_Analitico.py

echo.
echo Aguardando 6 minutos (360s) para o portal processar o Nao Visitadas...
timeout /t 360 /nobreak

echo [PASSO 2] Solicitando 04_Nao_Visitadas...
python 04_Nao_Visitadas.py

echo.
echo Aguardando 4 minutos (240s) para o portal processar o N...
timeout /t 240 /nobreak

echo.
echo [PASSO 3] Iniciando o Robo Coletor...
python 00_Coletor_FDS.py

echo.
echo ======================================================
echo   PROCESSO SIMPLES FINALIZADO!
echo ======================================================
exit