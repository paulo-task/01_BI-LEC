@echo off
cls
echo ======================================================
echo   SOLICITACAO SIMPLES (PDL PRIMEIRO, DEPOIS ELF)
echo ======================================================

echo [PASSO 1] Solicitando 02_PDL_Analitico...
:: Removi o 'start' para o log aparecer aqui, se preferir nova janela, volte o start
python 02_PDL_Analitico.py

echo.
echo Aguardando 3 minutos (180s) para o portal processar o PDL...
timeout /t 180 /nobreak

echo.
echo [PASSO 2] Solicitando 04_Nao_Visitadas...
:: Removi o 'start' para o log aparecer aqui, se preferir nova janela, volte o start
python 04_Nao_visitadas.py

echo.
echo Aguardando 3 minutos (180s) para o portal processar o PDL...
timeout /t 180 /nobreak

echo.
echo [PASSO 3] Solicitando 01_ELF_Hora...
python 01_ELF_Hora.py

echo.
echo Aguardando mais 5 minutos para o processamento do ELF...
timeout /t 300 /nobreak

echo.
echo [PASSO 3] Iniciando o Robo Coletor...
:: IMPORTANTE: Verifique se o nome do arquivo abaixo e exatamente como voce salvou
:: Use aspas se houver espaços no nome do arquivo .py
python 00_Coletor.py "Leitura" "Visitadas" "Produtividade" 

echo.
echo ======================================================
echo   PROCESSO SIMPLES FINALIZADO!
echo ======================================================
exit