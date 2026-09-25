@echo off
rem Abre o Escanner em Lote usando o Python instalado neste computador.
where pythonw >nul 2>nul && (start "" pythonw "%~dp0escanner.py" & exit /b)
where pyw >nul 2>nul && (start "" pyw "%~dp0escanner.py" & exit /b)
echo Python nao encontrado. Rode primeiro o arquivo instalar.bat.
pause
