@echo off
setlocal enabledelayedexpansion

set "R_VERSION=4.6.1"
set "R_INSTALLER=R-4.6.1-win.exe"
set "INSTALL_SCRIPT=instalar_entorno_verAC_v3.10.R"

title verAC - Instalador de entorno (IMPaCT)

rem Este .bat vive en la subcarpeta 'instalador'. La carpeta del programa es la de arriba.
pushd "%~dp0.."
set "BASE_DIR=%CD%\"
popd
set "SCRIPTS_DIR=%BASE_DIR%scripts"
set "PKG_DIR=%SCRIPTS_DIR%\paquetesR"
set "INSTALL_PATH=%SCRIPTS_DIR%\%INSTALL_SCRIPT%"
set "R_HOME=%SCRIPTS_DIR%\R\R-%R_VERSION%"
set "R_LOCAL_RSCRIPT=%R_HOME%\bin\x64\Rscript.exe"

echo(
echo ==================================================================
echo     verAC - INSTALADOR DE ENTORNO (IMPaCT)
echo ==================================================================
echo(
echo Directorio: %BASE_DIR%
echo(

rem --- Comprobacion de 64 bits ---
if /I not "%PROCESSOR_ARCHITECTURE%"=="AMD64" if not defined PROCESSOR_ARCHITEW6432 (
    echo ERROR: Este programa requiere Windows de 64 bits.
    echo Su equipo parece ser de 32 bits, y R 4.x no es compatible. Contacte con informatica.
    goto :fin
)

if not exist "%SCRIPTS_DIR%\" (
    echo ERROR: No existe la carpeta 'scripts' en: %BASE_DIR%
    goto :fin
)
if not exist "%INSTALL_PATH%" (
    echo ERROR: No se encuentra: %INSTALL_PATH%
    goto :fin
)

rem --- Instalar R en local si no esta ya desplegado ---
if exist "%R_LOCAL_RSCRIPT%" (
    echo R ya esta desplegado en: %R_HOME%
    goto :have_r
)
if exist "%R_HOME%\bin\Rscript.exe" (
    set "R_LOCAL_RSCRIPT=%R_HOME%\bin\Rscript.exe"
    echo R ya esta desplegado en: %R_HOME%
    goto :have_r
)

echo R no encontrado. Instalando R %R_VERSION% (local, sin permisos de administrador)...
if not exist "%PKG_DIR%\%R_INSTALLER%" (
    echo ERROR: no se encuentra el instalador de R en: %PKG_DIR%\%R_INSTALLER%
    goto :fin
)
echo Esto puede tardar uno o dos minutos. Espere, por favor...
start /wait "" "%PKG_DIR%\%R_INSTALLER%" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /DIR="%R_HOME%" /COMPONENTS="main,x64"

set "R_LOCAL_RSCRIPT=%R_HOME%\bin\x64\Rscript.exe"
for /L %%i in (1,1,10) do (
    if exist "%R_HOME%\bin\x64\Rscript.exe" set "R_LOCAL_RSCRIPT=%R_HOME%\bin\x64\Rscript.exe" ^& goto :r_ok
    if exist "%R_HOME%\bin\Rscript.exe" set "R_LOCAL_RSCRIPT=%R_HOME%\bin\Rscript.exe" ^& goto :r_ok
    timeout /t 3 /nobreak >nul
)
:r_ok
if not exist "%R_LOCAL_RSCRIPT%" (
    echo ERROR: la instalacion de R no se completo. Solicite a informatica instalar R en: %R_HOME%
    goto :fin
)
echo R %R_VERSION% instalado correctamente.

:have_r
echo(
echo Instalando paquetes y comprobando el entorno...
echo --------------------------------------------------------------------------------
"%R_LOCAL_RSCRIPT%" "%INSTALL_PATH%"
set "EXITCODE=!ERRORLEVEL!"
echo --------------------------------------------------------------------------------

if "!EXITCODE!"=="0" (
    echo Resultado: instalacion completada y verificada correctamente.
) else if "!EXITCODE!"=="3" (
    echo Resultado: instalado, pero SIN verificar. Revise el aviso de arriba.
) else if "!EXITCODE!"=="4" (
    echo Resultado: instalado, pero la verificacion FALLO. Revise el aviso de arriba.
) else (
    echo Resultado: la instalacion FALLO. Revise los mensajes de arriba.
    echo            Contacto: veronica.cabanas@uam.es
)

rem --- Crear acceso directo (.lnk) en la carpeta del programa, con icono ---
rem Se crea siempre, para que el tecnico tenga el acceso aunque la instalacion diera algun
rem aviso. Apunta al programa 'scripts\verAC_v3.10.bat' y usa el icono de 'scripts\logo\'.
set "LNK_PATH=%BASE_DIR%verAC.lnk"
set "TARGET_BAT=%SCRIPTS_DIR%\verAC_v3.10.bat"
set "ICON_PATH=%SCRIPTS_DIR%\logo\verAC_logo_horizontal.ico"
if not exist "%TARGET_BAT%" goto :fin
echo(
echo Creando acceso directo 'verAC' en la carpeta del programa...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%LNK_PATH%'); $s.TargetPath='%TARGET_BAT%'; $s.WorkingDirectory='%SCRIPTS_DIR%'; if (Test-Path '%ICON_PATH%') { $s.IconLocation='%ICON_PATH%' }; $s.Save()" >nul 2>&1
if exist "%LNK_PATH%" echo Acceso directo creado: %LNK_PATH%
if not exist "%LNK_PATH%" echo Aviso: no se pudo crear el acceso directo automaticamente. Puede crearlo a mano.

:fin
echo(
echo ================================================================================
echo Proceso finalizado.
echo ================================================================================
echo(
echo  Revise los mensajes de arriba. Cuando termine de leerlos, pulse una tecla
echo  o cierre la ventana con la X para salir.
echo(
pause >nul
endlocal
exit /b
