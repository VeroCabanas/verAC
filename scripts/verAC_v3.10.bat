@echo off
setlocal enabledelayedexpansion

set "SCRIPT_NAME=verAC_script_VCS_v3.10.R"
set "R_VERSION=4.6.1"

title verAC - Procesador de Archivos BIN (IMPaCT)

rem Este .bat vive DENTRO de la carpeta 'scripts'. La carpeta del programa es la de arriba.
rem %~dp0 = ...\verAC\scripts\   ;  BASE_DIR (raiz del programa) = un nivel por encima.
set "SCRIPTS_DIR=%~dp0"
pushd "%~dp0.."
set "BASE_DIR=%CD%\"
popd
set "SCRIPT_PATH=%SCRIPTS_DIR%%SCRIPT_NAME%"
set "R_HOME=%SCRIPTS_DIR%R\R-%R_VERSION%"
set "R_LOCAL_RSCRIPT=%R_HOME%\bin\x64\Rscript.exe"
if not exist "%R_LOCAL_RSCRIPT%" set "R_LOCAL_RSCRIPT=%R_HOME%\bin\Rscript.exe"
if exist "%R_LOCAL_RSCRIPT%" (
    set "R_RUNNER=%R_LOCAL_RSCRIPT%"
) else (
    for /f "delims=" %%I in ('where Rscript.exe 2^>nul') do if not defined R_RUNNER set "R_RUNNER=%%I"
)

echo(
echo ==================================================================
echo     PRE-PROCESADO DE ARCHIVOS .BIN DE ACELEROMETRIA - IMPaCT
echo ==================================================================
echo(
echo Directorio de trabajo: %BASE_DIR%
echo(

rem --- Comprobacion de 64 bits ---
if /I not "%PROCESSOR_ARCHITECTURE%"=="AMD64" if not defined PROCESSOR_ARCHITEW6432 (
    echo ERROR: Este programa requiere Windows de 64 bits. Contacte con informatica.
    goto :fin
)

if not defined R_RUNNER (
    echo ================================================================================
    echo  ATENCION: el entorno de verAC todavia no esta instalado en este equipo.
    echo ================================================================================
    echo(
    echo  No se ha encontrado R en la carpeta del programa ni en el sistema, asi que
    echo  verAC no puede procesar archivos todavia.
    echo(
    echo  QUE HACER:
    echo(
    echo   1. Instale R 4.4 o posterior y ejecute
    echo      'Rscript scripts\instalar_dependencias.R' con conexion a Internet;
    echo      o prepare el paquete autonomo y ejecute 'Instalador_verAC.bat'.
    echo(
    echo   2. Si el problema continua, ejecute 'Instalador_verAC.bat' desde una cuenta
    echo      con permisos de administrador, y compruebe que el antivirus o el cortafuegos
    echo      no esten bloqueando la instalacion. Si no esta segura, pida al servicio de
    echo      informatica de su centro que lo ejecute.
    echo(
    echo   3. Si aun asi no funciona, pongase en contacto con la desarrolladora del
    echo      programa: veronica.cabanas@uam.es
    echo(
    echo ================================================================================
    goto :fin
)
if /I "%R_RUNNER%"=="%R_LOCAL_RSCRIPT%" echo Usando R propio de verAC: %R_HOME%
if /I not "%R_RUNNER%"=="%R_LOCAL_RSCRIPT%" echo Usando R del sistema: %R_RUNNER%

if not exist "%SCRIPT_PATH%" (
    echo ERROR: No se encuentra el script principal:
    echo    %SCRIPT_PATH%
    echo Si el problema continua, contacte con: veronica.cabanas@uam.es
    goto :fin
)

echo(
echo INICIANDO PROCESAMIENTO...
echo Hora de inicio: %date% %time%
echo ================================================================================
pushd "%BASE_DIR%"
"%R_RUNNER%" "%SCRIPT_PATH%"
set "EXITCODE=!ERRORLEVEL!"
popd

if "!EXITCODE!"=="0" (
    echo(
    echo Hora de finalizacion: %date% %time%
    echo ================================================================================
    goto :fin
)

echo(
echo ================================================================================
echo PROCESO TERMINADO CON ERRORES ^(codigo !EXITCODE!^)
echo ================================================================================
if "!EXITCODE!"=="1" (
    echo Codigo 1: No se encontraron archivos .bin en 'archivos bin'.
) else if "!EXITCODE!"=="2" (
    echo Codigo 2: Problema con paquetes/entorno R.
    echo    Ejecute 'Rscript scripts\instalar_dependencias.R' o reinstale el entorno autonomo.
) else if "!EXITCODE!"=="3" (
    echo Codigo 3: Archivo abierto que debe sobrescribirse.
    echo    Cierre el Excel de resultados y los informes HTML/PDF y reejecute.
) else (
    echo Codigo !EXITCODE!: Error durante el procesamiento.
)
echo(
echo Contacto: veronica.cabanas@uam.es

:fin
echo(
echo ================================================================================
echo Proceso finalizado. Pulse una tecla o cierre la ventana con la X para salir.
pause >nul
endlocal
exit /b
