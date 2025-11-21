@echo off
REM Define o local atual do script .bat como a pasta de trabalho
cd /d "%~dp0"

echo ------------------------------------------
echo INICIANDO OFERTEIRO 
echo ------------------------------------------

REM Tenta executar o script juntarpdf.py usando o comando 'python'
python juntarpdf.py

REM Verifica o código de retorno do comando anterior
if errorlevel 1 (
    echo.
    echo ERRO: Ocorreu um erro ao rodar o script Python.
    echo Verifique o traceback acima ou se 'python' esta no PATH.
) else (
    echo.
    echo SUCESSO: O script juntarpdf.py terminou a execucao.
)

echo.
pause