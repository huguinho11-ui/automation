@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1"
if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar a automacao.
    pause
)
endlocal
