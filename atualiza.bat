@echo off
colort
echo Processando os novos dados da planilha...
python process_dashboard_data.py
if %errorlevel% neq 0 (
    echo Erro ao processar os dados! Verifique se a planilha esta fechada e com os nomes de colunas corretos.
    pause
    exit /b %errorlevel%
)

echo.
echo ====== Tudo pronto localmente! ======
echo.
set /p confirm="Deseja enviar agora a nova atualizacao para o site no Github? (S/N): "
if /i "%confirm%"=="S" (
    echo.
    echo Atualizando o Github...
    git add .
    git commit -m "Atualizando dados do dashboard"
    git push
    echo.
    echo Sucesso! O novo dashboard estara no ar em ate 5 minutinhos.
    pause
) else (
    echo.
    echo Atualizacao abortada pelo usuario.
    pause
)
