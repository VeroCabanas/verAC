# verAC

**verAC** (*verificación de archivos de ACelerometría*) es una herramienta para Windows que automatiza la lectura, la comprobación inicial y la organización de archivos `.bin` generados por dispositivos Matrix. Fue desarrollada para el control de calidad de los registros de acelerometría de muñeca y muslo de la cohorte IMPaCT.

La aplicación utiliza R y [`GGIRread`](https://cran.r-project.org/package=GGIRread) para decodificar los archivos y produce, para cada registro:

- clasificación operativa del archivo como válido, (semi) válido o corrupto;
- fechas y duración del registro;
- frecuencia de muestreo y rango dinámico del acelerómetro;
- identificación de los sensores disponibles;
- un informe individual en HTML;
- un Excel individual y un Excel global de control;
- una estructura ordenada por código de participante.

## Dos formas de distribución

Este repositorio contiene el **código fuente, los lanzadores, la documentación y los metadatos**. No incorpora el entorno completo de R, los paquetes binarios ni datos de participantes.

La futura *release* incluirá, como archivo independiente, un **paquete autónomo para Windows** preparado para funcionar sin conexión a Internet y, en la mayoría de los equipos, sin permisos de administración. Mantener ese paquete fuera del historial de Git hace que el repositorio sea ligero y permite documentar correctamente los componentes de terceros.

## Requisitos

- Windows de 64 bits.
- Para el paquete autónomo: Windows 10, Windows Server 2016 o posterior. R 4.6.1 requiere UCRT.
- Archivos Matrix/Parmay `.bin` compatibles con `GGIRread`.
- Nombres con `ACEMS` para muñeca y `ACEMI` para muslo cuando se quiera identificar automáticamente la posición.

## Uso básico

1. Copiar los archivos `.bin` en `archivos bin`.
2. Ejecutar `verAC_v3.10.bat` desde el acceso directo o desde `scripts`.
3. Consultar:
   - `Excel de resultados del procesado/RESULTADOS_COMPROBACION_DE_ARCHIVOS_ALL.xlsx`;
   - `archivos bin procesados/<código>/` para los informes y resultados individuales.

Las instrucciones completas están en [docs/USO.md](docs/USO.md). La instalación del paquete autónomo se describe en [docs/INSTALACION.md](docs/INSTALACION.md).

> [!IMPORTANT]
> verAC **mueve** los `.bin` procesados desde `archivos bin` a la carpeta del participante. Cuando se reprocesa un archivo cuyo nombre ya existe y las copias de seguridad están desactivadas, el archivo de destino puede sustituirse. Antes de trabajar con datos irremplazables, conserve una copia maestra fuera de la carpeta de verAC.

## Protección de datos

No se incluyen datos reales en este repositorio. Los nombres de archivo y los resultados pueden contener códigos de participante y metadatos temporales, por lo que deben tratarse conforme al protocolo del estudio y a la normativa aplicable. Véase [docs/PROTECCION_DE_DATOS.md](docs/PROTECCION_DE_DATOS.md).

## Alcance y limitaciones

- verAC realiza un **control de calidad técnico inicial**; no sustituye la revisión científica del registro.
- La etiqueta «corrupto» refleja que el archivo no pudo decodificarse con la estructura esperada. Los errores del entorno se distinguen y no se clasifican como corrupción del dato.
- No es un producto sanitario ni está destinado a apoyar decisiones clínicas.
- La versión 3.10 se ha validado en Windows de 64 bits con R 4.6.1 y `GGIRread` 1.0.8.
- Los dos registros breves incluidos en `scripts/arch_prueba` fueron creados expresamente para comprobar el funcionamiento del entorno y están autorizados para su distribución.

## Código, dependencias y licencia

El código propio de verAC se distribuye con licencia [MIT](LICENSE). R y los paquetes incluidos en el futuro paquete autónomo conservan sus respectivas licencias. La relación se documenta en [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) y [dependency-manifest.csv](dependency-manifest.csv).

## Cita

Los metadatos de citación están disponibles en [`CITATION.cff`](CITATION.cff). Cuando se publique una versión estable y se deposite en Zenodo, se añadirá aquí su DOI.

## Autoría

Verónica Cabanas Sánchez<br>
Universidad Autónoma de Madrid<br>
[ORCID 0000-0003-1235-3535](https://orcid.org/0000-0003-1235-3535)
