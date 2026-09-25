@echo off
title Instalar Escanner em Lote
echo ============================================
echo   Instalando o Escanner em Lote
echo ============================================
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Python nao encontrado.
  echo Instale o Python pelo site python.org e marque a opcao
  echo "Add python.exe to PATH" na primeira tela do instalador.
  echo Depois rode este arquivo de novo.
  pause
  exit /b 1
)
echo [1/2] Instalando componentes (pode demorar alguns minutos)...
python -m pip install --disable-pip-version-check -q -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo [ERRO] Falha ao instalar os componentes. Verifique a internet e tente de novo.
  pause
  exit /b 1
)
echo [2/2] Criando atalho na Area de Trabalho...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar_atalho.ps1"
if errorlevel 1 (
  echo [AVISO] Nao consegui criar o atalho. Use o arquivo "Abrir Escanner.bat".
) else (
  echo Atalho "Escanner em Lote" criado na Area de Trabalho.
)
echo.
echo Pronto! Abra o programa pelo atalho "Escanner em Lote".
pause
