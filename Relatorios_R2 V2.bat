@echo off
cls
echo ======================================================
echo   INICIANDO CICLO COMPLETO (SOLICITACAO + COLETA)
echo ======================================================
echo.

echo [PASSO 1] Disparando 1º Bloco de Solicitacoes (02, 04)...
start python 02_PDL_Analitico.py
start python 04_Nao_Visitadas.py

echo Aguardando 90 segundos para estabilidade do portal...
timeout /t 90 /nobreak

echo [PASSO 2] Disparando 2º Bloco de Solicitacoes (05, 06, 07)...
start python 05_Nao_Visitadas_Historico.py
start python 06_Impedimentos.py
start python 07_Entregas.py

echo Aguardando 300 segundos para estabilidade do portal...
timeout /t 300 /nobreak

echo [PASSO 3] Disparando 3º Bloco de Solicitacoes (01, 03)...
start python 01_ELF_Hora.py
start python 03_Nao_Lib_Fatura.py

echo.
echo ======================================================
echo   TODOS OS PEDIDOS FORAM ENVIADOS AO PORTAL!
echo   Aguardando 8 minutos (480s) para o processamento...
echo ======================================================
timeout /t 480 /nobreak

echo.
echo [PASSO 3] Iniciando o Robo Coletor para Baixar e Organizar...
:: Note que aqui passamos o argumento "completo", para ele baixar os 7 relatorios
python 00a_Coletor.py completo

echo.
echo ======================================================
echo   PROCESSO FINALIZADO COM SUCESSO!
echo ======================================================
pause
exit