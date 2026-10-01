@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel% equ 0 (
  py -3 Relatorio_saidas.py
) else (
  python Relatorio_saidas.py
)
if errorlevel 1 (
  echo.
  echo Instale Python 3.10 ou superior e marque Add Python to PATH.
  pause
)
