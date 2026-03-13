@echo off
title Extractor de Transferencias de WhatsApp
color 0A

echo ================================================================
echo       EXTRACTOR DE TRANSFERENCIAS - RED POSTAL POBLADO
echo ================================================================
echo.

:: Verificar si Node.js esta instalado
node -v >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js no esta instalado o no fue configurado correctamente.
    echo Por favor descarga e instala Node.js (la version LTS) desde:
    echo https://nodejs.org/es
    echo.
    echo Una vez instalado, vuelve a abrir este archivo.
    pause
    exit /b
)

:: Revisar dependencias e instalarlas si faltan
IF NOT EXIST "node_modules\" (
    echo [INFO] Es la primera vez que se ejecuta en esta computadora.
    echo Instalando dependencias necesarias (esto puede tardar unos minutos)...
    call npm install
    echo.
)

:: Ejecutar el script principal
echo [INFO] Iniciando el sistema...
echo.
call node index.js

echo.
echo ================================================================
echo   LA HERRAMIENTA HA FINALIZADO SU EJECUCION
echo ================================================================
pause
