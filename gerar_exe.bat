@echo off
title Gerar Escanner em Lote.exe
rem Recria o "Escanner em Lote.exe" a partir dos arquivos .py desta pasta
rem e copia o novo exe para a Area de Trabalho.
cd /d "%~dp0"
python -m pip install --disable-pip-version-check -q pyinstaller -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name "Escanner em Lote" --icon "%~dp0icone.ico" --add-data "%~dp0icone.ico;." --collect-all winrt --hidden-import win32timezone --distpath dist --workpath "%TEMP%\escanner_build" --specpath "%TEMP%\escanner_build" escanner.py
if errorlevel 1 (
  echo [ERRO] Nao foi possivel gerar o exe.
  pause
  exit /b 1
)
for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do copy /y "dist\Escanner em Lote.exe" "%%d\Escanner em Lote.exe" >nul
echo.
echo Pronto! "Escanner em Lote.exe" atualizado na Area de Trabalho.
pause
