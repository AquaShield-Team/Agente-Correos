@echo off
chcp 65001 >nul 2>&1
title AquaShield · Instalador Agente de Correos
color 0B

echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║   AquaShield · Instalador Agente de Correos             ║
echo  ║   Configuracion automatica del entorno                  ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: ── 1. Verificar Python ──────────────────────────────────────
echo [1/3] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ❌ Python no esta instalado o no esta en el PATH.
    echo     Descarga Python 3.11+ desde: https://www.python.org/downloads/
    echo     IMPORTANTE: Marca la casilla "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo     ✅ Python %PYVER% detectado.

:: ── 2. Verificar Git ─────────────────────────────────────────
echo.
echo [2/3] Verificando Git...
git --version >nul 2>&1
if errorlevel 1 (
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
    pause
    exit /b 0
)
for /f "tokens=3" %%v in ('git --version 2^>^&1') do echo     ✅ Git %%v detectado.

:: ── 3. Instalar dependencias Python ──────────────────────────
echo.
echo [3/3] Instalando dependencias de Python...
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
echo  ╚══════════════════════════════════════════════════════════╝
echo.
pause
