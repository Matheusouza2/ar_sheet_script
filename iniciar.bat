@echo off
title Gerador de Ordem de Servico
cd /d "%~dp0"

set "PORTABLE_DIR=%CD%\_portable"
set "PYTHON_DIR=%PORTABLE_DIR%\python"
set "PYTHON_VER=3.12.4"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VER%/python-%PYTHON_VER%-amd64.exe"

echo =============================================
echo   Gerador de Ordem de Servico
echo =============================================
echo.

:: ---------------------------------------------------------
:: PASSO 1 — Encontrar ou baixar Python
:: ---------------------------------------------------------
set "PYTHON_EXE="

:: Tenta Python global
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_EXE=python"
    goto :instalar_deps
)

:: Tenta Python portatil ja baixado
if exist "%PYTHON_DIR%\python.exe" (
    set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
    goto :instalar_deps
)

:: Baixa e instala Python portatil
echo [PASSO 1/4] Python nao encontrado no sistema.
echo Baixando Python %PYTHON_VER% (portatil)...
if not exist "%PORTABLE_DIR%" mkdir "%PORTABLE_DIR%"

:: Tenta curl (Win10+), fallback para PowerShell
where curl >nul 2>&1
if %errorlevel% equ 0 (
    curl -# -L -o "%PORTABLE_DIR%\installer.exe" "%PYTHON_URL%"
) else (
    powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PORTABLE_DIR%\installer.exe'"
)

if %errorlevel% neq 0 (
    echo [ERRO] Falha ao baixar Python. Verifique sua conexao de internet.
    pause
    exit /b 1
)

echo [PASSO 2/4] Instalando Python em diretorio local...
"%PORTABLE_DIR%\installer.exe" /quiet ^
    InstallAllUsers=0 ^
    TargetDir="%PYTHON_DIR%" ^
    Include_test=0 ^
    Include_launcher=0 ^
    Include_pip=1 ^
    Include_tcltk=1

del "%PORTABLE_DIR%\installer.exe"

if not exist "%PYTHON_DIR%\python.exe" (
    echo [ERRO] Falha ao instalar Python.
    pause
    exit /b 1
)

set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
echo [OK] Python %PYTHON_VER% instalado em %PYTHON_DIR%

:: ---------------------------------------------------------
:: PASSO 3 — Instalar dependencias
:: ---------------------------------------------------------
:instalar_deps
echo [PASSO 3/4] Instalando dependencias...
"%PYTHON_EXE%" -m pip install --upgrade pip --quiet
"%PYTHON_EXE%" -m pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.

:: ---------------------------------------------------------
:: PASSO 4 — Executar aplicacao
:: ---------------------------------------------------------
:executar
echo [PASSO 4/4] Iniciando aplicacao...
echo.
"%PYTHON_EXE%" main.py
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] A aplicacao encerrou com erro.
    pause
)
exit /b 0
