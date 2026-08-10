# Uso de verAC

## Preparación

1. Mantenga una copia maestra de los datos originales fuera de verAC.
2. Copie en `archivos bin` los registros Matrix que quiera comprobar.
3. Utilice nombres que incluyan:
   - `ACEMS` para registros de muñeca;
   - `ACEMI` para registros de muslo;
   - un guion antes del código del participante, por ejemplo `MD ACEMS1-12345678.bin`.

Si el nombre no sigue este patrón, verAC puede clasificar la posición o el código como desconocidos aunque el archivo sea legible.

## Ejecución

Ejecute `verAC.lnk` o `scripts/verAC_v3.10.bat`. El programa procesa todos los `.bin` presentes en la carpeta de entrada.

Para cada archivo, verAC:

1. intenta decodificar el registro con `GGIRread`;
2. comprueba la estructura básica de acelerometría;
3. resume duración, frecuencia, rango y sensores;
4. genera el informe HTML y los registros Excel;
5. mueve el `.bin` a la carpeta correspondiente del participante.

## Resultados

- `Excel de resultados del procesado/RESULTADOS_COMPROBACION_DE_ARCHIVOS_ALL.xlsx`: inventario global.
- `archivos bin procesados/<código>/Resultados_<código>.xlsx`: resultado individual.
- `archivos bin procesados/<código>/informe_<archivo>.html`: informe visual.
- `archivos bin procesados/<código>/<archivo>.bin`: original reorganizado.

## Estados

- **válido**: el registro se ha decodificado y contiene la estructura de acelerometría esperada.
- **(semi) válido**: el registro es utilizable, pero presenta truncamiento o paquetes imputados según los criterios implementados.
- **corrupto**: no se ha podido extraer una estructura de datos válida del archivo.
- **omitido por error de entorno**: verAC ha detectado un problema de R, de una dependencia o de escritura; el archivo permanece en la entrada para volver a intentarlo.

## Reprocesamiento y copias de seguridad

En la configuración original de la versión 3.10:

- `reprocesar_existentes = TRUE`;
- las copias de seguridad de `.bin`, HTML y Excel están desactivadas;
- el Excel individual elimina la fila anterior del mismo archivo;
- el Excel global mantiene el historial.

Revise el bloque `CONFIG` al principio de `scripts/verAC_script_VCS_v3.10.R` antes de cambiar estas opciones.
