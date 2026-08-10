# Arquitectura del paquete autónomo

La distribución sin conexión se construye superponiendo al contenido del repositorio los siguientes componentes, que no se guardan en Git:

```text
verAC/
├── archivos bin/
├── archivos bin procesados/
├── Excel de resultados del procesado/
├── instalador/
│   └── Instalador_verAC.bat
└── scripts/
    ├── R/R-4.6.1/
    ├── paquetesR/
    │   ├── R-4.6.1-win.exe
    │   └── paquetes binarios .zip
    ├── arch_prueba/
    ├── logo/
    ├── instalar_entorno_verAC_v3.10.R
    ├── verAC_script_VCS_v3.10.R
    └── verAC_v3.10.bat
```

## Motivo de la separación

- Evita versionar miles de archivos binarios y un entorno de cientos de megabytes.
- Impide que los datos reales entren accidentalmente en el historial.
- Permite publicar el paquete autónomo como activo versionado de una *release*.
- Facilita acompañar cada distribución con manifiestos, sumas de comprobación, licencias y fuentes correspondientes.

## Verificaciones antes de una release

1. Confirmar la procedencia y la posibilidad de redistribuir los `.BIN` de autocomprobación.
2. Generar SHA-256 de todos los binarios incluidos.
3. Comprobar que el instalador de R coincide con la huella oficial de CRAN.
4. Ejecutar la autocomprobación en una carpeta limpia.
5. Ejecutar un ciclo completo sobre datos de prueba y revisar HTML y Excel.
6. Examinar el paquete para confirmar que no contiene códigos de participantes ni resultados reales.
7. Adjuntar las licencias y las fuentes correspondientes de los componentes redistribuidos.
