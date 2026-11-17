@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Motor de estados financieros NIIF

set "PY="
set "VENV_PY=%~dp0.venv\Scripts\python.exe"

rem ---- 1. Si el entorno ya esta listo, ir directo al programa ----
if exist "%VENV_PY%" if exist "%~dp0.venv\instalado.ok" goto ejecutar

rem ---- 2. Buscar un Python real (no el acceso directo de Microsoft Store) ----
:buscar_python
set "PY="
py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PY="%%D\python.exe""
if not defined PY for /d %%D in ("%ProgramFiles%\Python3*") do if exist "%%D\python.exe" set "PY="%%D\python.exe""
if defined PY goto crear_entorno

echo.
echo ==============================================================
echo  No se encontro Python en este computador.
echo  El programa necesita Python 3.10 o superior para funcionar.
echo ==============================================================
echo.
set "RESP="
set /p "RESP=Desea instalar Python ahora? Escriba S y presione Enter: "
if /i not "%RESP%"=="S" goto instalar_manual

where winget >nul 2>&1
if errorlevel 1 goto instalar_manual

echo.
echo Instalando Python 3.12. Si Windows pide permiso o pide aceptar
echo los terminos, responda que si. Puede tardar unos minutos...
echo.
winget install -e --id Python.Python.3.12 --scope user
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
    goto crear_entorno
)
goto buscar_python_despues

:buscar_python_despues
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PY="%%D\python.exe""
if defined PY goto crear_entorno
goto instalar_manual

:instalar_manual
echo.
echo Se abrira la pagina de descarga de Python.
echo   1. Descargue el instalador y abralo.
echo   2. IMPORTANTE: marque la casilla "Add python.exe to PATH".
echo   3. Presione "Install Now" y espere a que termine.
echo   4. Vuelva a abrir ejecutar.bat.
echo.
start "" "https://www.python.org/downloads/windows/"
pause
exit /b 1

rem ---- 3. Crear el entorno virtual e instalar las librerias ----
:crear_entorno
if exist "%~dp0.venv" rmdir /s /q "%~dp0.venv"
echo.
echo Preparando el programa por primera vez (solo se hace una vez)...
%PY% -m venv .venv
if not exist "%VENV_PY%" (
    echo.
    echo No se pudo crear el entorno virtual con %PY%.
    pause
    exit /b 1
)
"%VENV_PY%" -m pip install --upgrade pip
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Fallo la instalacion de librerias. Revise la conexion a internet
    echo y vuelva a abrir ejecutar.bat.
    pause
    exit /b 1
)
echo ok> "%~dp0.venv\instalado.ok"

rem ---- 4. Ejecutar el motor ----
:ejecutar
echo.
"%VENV_PY%" main.py
if errorlevel 2 (
    echo.
    echo No se genero el Excel. Revise el mensaje de arriba.
) else (
    echo.
    echo El Excel quedo en la carpeta data\output
)
pause
