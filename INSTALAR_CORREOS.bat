@echo off
chcp 65001 >nul 2>&1
title AquaShield · Instalador Agente de Correos
color 0B

echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║   AquaShield · Instalador Agente de Correos             ║
echo  ║   Este script instala TODO lo necesario                  ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: ── 1. Verificar/Instalar Python ─────────────────────────────
echo [1/4] Verificando Python...
python --version >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=2" %%v in ('python --version 2^>^&1') do echo     ✅ Python %%v detectado.
) else (
    echo     ❌ Python no encontrado. Instalando con winget...
    winget install Python.Python.3.13 --accept-package-agreements --accept-source-agreements -e
    if errorlevel 1 (
        echo.
        echo  ⚠️  No se pudo instalar Python automaticamente.
        echo     Descarga manualmente: https://www.python.org/downloads/
        echo     IMPORTANTE: Marca "Add Python to PATH" al instalar.
        echo.
        pause
        exit /b 1
    )
    echo     ✅ Python instalado. Reinicia este script para continuar.
    echo     ⚠️  CIERRA esta ventana, abre CMD de nuevo y ejecuta este script otra vez.
    pause
    exit /b 0
)

:: ── 2. Verificar/Instalar Git ────────────────────────────────
echo.
echo [2/4] Verificando Git...
git --version >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=3" %%v in ('git --version 2^>^&1') do echo     ✅ Git %%v detectado.
) else (
    echo     ❌ Git no encontrado. Instalando con winget...
    winget install Git.Git --accept-package-agreements --accept-source-agreements -e
    if errorlevel 1 (
        echo.
        echo  ⚠️  No se pudo instalar Git automaticamente.
        echo     Descarga manualmente: https://git-scm.com/downloads/win
        echo.
        pause
        exit /b 1
    )
    echo     ✅ Git instalado. Reinicia este script para continuar.
    echo     ⚠️  CIERRA esta ventana, abre CMD de nuevo y ejecuta este script otra vez.
    pause
    exit /b 0
)

:: ── 3. Clonar repositorio (si no existe) ─────────────────────
echo.
echo [3/4] Verificando repositorio...
if exist "app.py" (
    echo     ✅ Repositorio ya descargado.
) else if exist "Agente-Correos\app.py" (
    echo     ✅ Repositorio ya clonado en subcarpeta.
    cd Agente-Correos
) else (
    echo     Descargando el Agente de Correos desde GitHub...
    git clone https://github.com/AquaShield-Team/Agente-Correos.git
    if errorlevel 1 (
        echo.
        echo  ⚠️  No se pudo descargar. Revisa tu acceso a GitHub.
        pause
        exit /b 1
    )
    cd Agente-Correos
    echo     ✅ Repositorio descargado.
)

:: ── 4. Instalar dependencias Python ──────────────────────────
echo.
echo [4/4] Instalando dependencias de Python...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo.
    echo  ⚠️  Algunas dependencias fallaron.
    echo     Si estas detras de un firewall, conecta via hotspot.
    pause
    exit /b 1
)
echo     ✅ Dependencias instaladas correctamente.

:: ── Crear carpetas necesarias ────────────────────────────────
if not exist "uploads" mkdir uploads

:: ── Crear Base_Clientes.xlsx de ejemplo si no existe ─────────
if not exist "Base_Clientes.xlsx" (
    echo.
    echo     📄 Creando Base_Clientes.xlsx de ejemplo...
    python -c "import openpyxl; wb=openpyxl.Workbook(); ws=wb.active; ws.append(['ID','Cliente','Para','CC','Asunto','Cuerpo']); ws.append(['1','Cliente Ejemplo','correo@ejemplo.com','','Documentos [CLIENTE]','Estimado,<br>Adjunto documentos.']); wb.save('Base_Clientes.xlsx'); print('     ✅ Base_Clientes.xlsx creada.')"
)

:: ── Listo ────────────────────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║   ✅ INSTALACION COMPLETADA                             ║
echo  ╠══════════════════════════════════════════════════════════╣
echo  ║                                                          ║
echo  ║   Para usar el agente:                                   ║
echo  ║   1. Ejecuta Iniciar_Agente_Correos.bat                  ║
echo  ║   2. Se abre el dashboard en tu navegador                ║
echo  ║   3. Edita Base_Clientes.xlsx para agregar clientes      ║
echo  ║   4. Selecciona un cliente y genera el correo!           ║
echo  ║                                                          ║
echo  ║   Las actualizaciones se descargan automaticamente       ║
echo  ║   cada vez que inicies el agente.                        ║
echo  ║                                                          ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.
pause
