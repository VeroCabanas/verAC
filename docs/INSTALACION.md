# Instalación

## Paquete autónomo para Windows

La distribución autónoma está pensada para equipos sin conexión a Internet y sin una instalación previa de R.

1. Descomprima la carpeta completa de verAC en una ubicación con permiso de escritura.
2. Abra `instalador` y ejecute `Instalador_verAC.bat`.
3. El instalador despliega R 4.6.1 dentro de `scripts/R/R-4.6.1`, instala las dependencias desde los `.zip` locales y ejecuta una autocomprobación.
4. Si la comprobación finaliza correctamente, se crea el acceso directo `verAC.lnk` en la raíz del programa.

No se modifica una instalación general de R ni se instala el entorno en `Program Files`. En la mayoría de los equipos no se necesitan permisos de administrador. Algunas políticas corporativas, antivirus o mecanismos de control de aplicaciones pueden impedir la ejecución de instaladores locales; en esos casos debe intervenir el servicio de informática.

## Compatibilidad

- Windows de 64 bits.
- Windows 10, Windows Server 2016 o posterior, salvo que se instale UCRT manualmente en sistemas anteriores.
- Espacio libre suficiente para el entorno, los archivos `.bin` y los resultados.

## Repositorio de código

El repositorio no contiene `scripts/R` ni `scripts/paquetesR`. Esas carpetas forman parte exclusivamente del paquete autónomo publicado como activo de una *release*.
