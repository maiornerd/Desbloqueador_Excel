@echo off
REM Script de instalação automática para Windows
REM Desbloqueador de Planilhas Excel

echo.
echo ============================================================
echo   DESBLOQUEADOR DE PLANILHAS EXCEL - Instalador Automático
echo ============================================================
echo.

REM Verificar se Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python não foi encontrado!
    echo.
    echo Por favor, instale Python 3.7+ em: https://www.python.org/downloads/
    echo Certifique-se de marcar "Add Python to PATH" durante a instalação
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado
echo.

REM Instalar dependências
echo [1/3] Instalando dependências necessárias...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar dependências
    pause
    exit /b 1
)

echo [OK] Dependências instaladas
echo.

REM Criar executável
echo [2/3] Criando executável...
python setup_build.py
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao criar executável
    pause
    exit /b 1
)

echo [OK] Executável criado
echo.

REM Informações finais
echo [3/3] Finalizando...
echo.
echo ============================================================
echo   SUCESSO! Instalação Concluída
echo ============================================================
echo.
echo [LOCAL DO ARQUIVO]
echo   dist\Desbloqueador Excel.exe
echo.
echo [PRÓXIMOS PASSOS]
echo   1. Abra a pasta "dist"
echo   2. Execute "Desbloqueador Excel.exe"
echo   3. Você pode copiar este arquivo para qualquer PC
echo.
echo [CRIAR ATALHO NO DESKTOP]
echo   1. Clique com botão direito no .exe
echo   2. Selecione "Enviar para" > "Desktop (criar atalho)"
echo.
pause
