# Instalación

## Opción A. R instalado en el equipo

Esta es la vía reproducible a partir del repositorio público.

1. Instale R 4.4 o posterior para Windows de 64 bits.
2. Descargue o clone el repositorio en una ubicación con permiso de escritura.
3. Con conexión a Internet, ejecute:

   ```text
   Rscript scripts/instalar_dependencias.R
   ```

4. Confirme que la instalación termina con las versiones requeridas de `GGIRread`, `data.table`, `openxlsx` y `zip`.
5. Ejecute `scripts/verAC_v3.10.bat`. El lanzador buscará `Rscript.exe` en el sistema cuando no encuentre el entorno portátil.

La generación opcional de PDF requiere además `webshot2` y Chrome o Chromium. El HTML se genera sin esa dependencia cuando `generar_pdf = FALSE`, que es el valor predeterminado.

## Opción B. Paquete autónomo para Windows

La distribución autónoma está pensada para equipos sin conexión a Internet y sin una instalación previa de R. No se distribuye públicamente junto con el repositorio debido a su tamaño y a los componentes binarios de terceros; puede reconstruirse de forma controlada siguiendo `PAQUETE_AUTONOMO.md`.

1. Descomprima la carpeta completa de verAC en una ubicación con permiso de escritura.
2. Abra `instalador` y ejecute `Instalador_verAC.bat`.
3. El instalador despliega R 4.6.1 dentro de `scripts/R/R-4.6.1`, instala las dependencias desde los `.zip` locales y ejecuta una autocomprobación.
4. Si la comprobación finaliza correctamente, se crea el acceso directo `verAC.lnk` en la raíz del programa.

No se modifica una instalación general de R ni se instala el entorno en `Program Files`. En la mayoría de los equipos no se necesitan permisos de administrador. Algunas políticas corporativas, antivirus o mecanismos de control de aplicaciones pueden impedir la ejecución de instaladores locales; en esos casos debe intervenir el servicio de informática.

## Compatibilidad

- Windows de 64 bits.
- Windows 10, Windows Server 2016 o posterior, salvo que se instale UCRT manualmente en sistemas anteriores.
- Espacio libre suficiente para el entorno, los archivos `.bin` y los resultados.

## Contenido del repositorio

El repositorio no contiene `scripts/R` ni `scripts/paquetesR`. Esas carpetas forman parte exclusivamente de las copias autónomas reconstruidas para entornos autorizados.
