@echo off
:: Detiene todos los procesos de Node.js en ejecución en la máquina local
taskkill /f /im node.exe >nul 2>&1
echo ===================================================
echo   Servidor local de MailMerge Studio detenido.
echo ===================================================
timeout /t 3 >nul
