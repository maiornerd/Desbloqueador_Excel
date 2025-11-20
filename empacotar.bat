@echo off
REM Script para copiar apenas o executável compilado
REM Útil para distribuição do programa

setlocal enabledelayedexpansion

echo.
echo ============================================================
echo   EMPACOTADOR DO DESBLOQUEADOR EXCEL
echo ============================================================
echo.

REM Verificar se o executável existe
if not exist "dist\Desbloqueador Excel.exe" (
    echo [ERRO] Executável não encontrado!
    echo.
    echo Por favor, compile o programa primeiro:
    echo   1. Execute: instalar.bat
    echo   OU
    echo   2. Execute: python setup_build.py
    echo.
    pause
    exit /b 1
)

echo [OK] Executável encontrado
echo.

REM Criar pasta de distribuição
if not exist "distribuicao" mkdir distribuicao

echo [1/3] Copiando executável...
copy "dist\Desbloqueador Excel.exe" "distribuicao\Desbloqueador Excel.exe" >nul
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao copiar executável
    pause
    exit /b 1
)

echo [OK] Executável copiado
echo.

echo [2/3] Criando arquivo de instruções...
(
    echo Desbloqueador de Planilhas Excel v1.0
    echo.
    echo INSTRUCOES DE USO:
    echo ==================
    echo.
    echo 1. Duplo clique em "Desbloqueador Excel.exe" para abrir
    echo 2. Carregue um arquivo Excel protegido
    echo 3. Clique em "Desbloquear Planilha"
    echo 4. Salve o arquivo desbloqueado
    echo.
    echo REQUISITOS:
    echo ============
    echo - Windows 7 ou superior
    echo - Nenhuma instalacao adicional
    echo.
    echo DUVIDAS?
    echo ========
    echo Consulte o README.md ou GUIA_INSTALACAO.md
    echo.
    echo Desenvolvido com Python 3.7^+
) > "distribuicao\LEIA-ME.txt"

echo [OK] Arquivo de instruções criado
echo.

echo [3/3] Finalizando...
echo.
echo ============================================================
echo   PACOTE DE DISTRIBUICAO PRONTO!
echo ============================================================
echo.
echo [ARQUIVO PRONTO PARA ENVIO]
echo   Pasta: distribuicao\
echo   Arquivo: Desbloqueador Excel.exe
echo   Tamanho: ~70-80 MB
echo.
echo [PROXIMOS PASSOS]
echo   1. Abra a pasta "distribuicao"
echo   2. Comprima em ZIP para enviar por email (tamanho ~25-30MB)
echo   3. Ou copie diretamente para uma unidade USB
echo.
echo [PARA OUTRO USUARIO]
echo   - Envie o arquivo "Desbloqueador Excel.exe"
echo   - Ele pode executar diretamente em Windows 7+
echo   - Nao precisa instalar nada!
echo.
pause
