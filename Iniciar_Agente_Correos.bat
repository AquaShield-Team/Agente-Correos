@echo off
chcp 65001 >nul 2>&1
title AquaShield - Agente de Correos
color 0B

echo.
echo   ====================================
echo     AquaShield - Agente de Correos
echo   ====================================
echo.

cd /d "%~dp0"

:: ── Auto-actualizar desde GitHub ─────────────────────────────
git --version >nul 2>&1
if %errorlevel%==0 (
    if exist ".git" (
        echo   Buscando actualizaciones...
        git pull origin main 2>nul
        if %errorlevel%==0 (
            echo   [OK] Actualizado a la ultima version.
            echo   Verificando dependencias...
            pip install -r requirements.txt -q 2>nul
            echo   [OK] Dependencias OK.
        ) else (
            echo   [!] Sin conexion - usando version local.
        )
        echo.
    )
)

:: ── Detener servidor anterior ────────────────────────────────
echo   Iniciando servidor...
echo.

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5055 " ^| findstr "LISTENING"') do (
    echo   Deteniendo servidor anterior [PID %%a]...
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 2 /nobreak >nul

:: Iniciar servidor en segundo plano
start /min "AquaShield-Correos" python app.py

:: Esperar a que el servidor este listo
echo   Esperando que el servidor responda...
:WAIT
timeout /t 1 /nobreak >nul
python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:5055/', timeout=2)" >nul 2>&1
if %errorlevel% neq 0 goto WAIT

echo   [OK] Servidor listo en http://127.0.0.1:5055
echo.

:: Abrir navegador
start "" "http://127.0.0.1:5055"

echo   Dashboard abierto en tu navegador.
echo   El servidor seguira corriendo en segundo plano.
echo.
echo   Para detenerlo, cierra la ventana "AquaShield-Correos"
echo.
timeout /t 5 /nobreak >nul
