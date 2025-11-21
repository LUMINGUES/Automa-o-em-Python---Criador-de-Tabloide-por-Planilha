@echo off
setlocal

echo =========================================================
echo  INICIANDO INSTALACAO DE DEPENDENCIAS PARA AUTODOC.PY
echo =========================================================

echo 1. Criando ambiente virtual (venv)...
python -m venv venv
if %errorlevel% neq 0 (
echo ERRO: Falha ao criar o ambiente virtual. Certifique-se de que o Python esta no PATH.
pause
exit /b 1
)

echo 2. Ativando ambiente virtual...
call venv\Scripts\activate
if %errorlevel% neq 0 (
echo ERRO: Falha ao ativar o ambiente virtual.
pause
exit /b 1
)

echo 3. Instalando pacotes necessarios (pandas, docx, requests, docx2pdf)...
python -m pip install pandas openpyxl python-docx requests docx2pdf
if %errorlevel% neq 0 (
echo ERRO: Falha ao instalar as dependencias. Tente executar o .bat como Administrador.
pause
exit /b 1
)

echo =========================================================
echo  INSTALACAO CONCLUIDA COM SUCESSO!
echo =========================================================

echo AGORA, voce precisa EXECUTAR o seu script DENTRO deste ambiente.

echo Para executar o autodoc.py, digite na linha de comando:
echo python autodoc.py

endlocal
pause