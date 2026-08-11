# verAC

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.21887036.svg)](https://doi.org/10.5281/zenodo.21887036)

**verAC** (*verificación de archivos de ACelerometría*) es una herramienta para Windows que automatiza la lectura, la comprobación inicial y la organización de archivos `.bin` generados por dispositivos Matrix. Fue desarrollada para el control de calidad de los registros de acelerometría de muñeca y muslo de la cohorte IMPaCT.


La aplicación utiliza R y [`GGIRread`](https://cran.r-project.org/package=GGIRread) para decodificar los archivos y produce, para cada registro:


- clasificación operativa del archivo como válido, (semi) válido o corrupto;
- fechas y duración del registro;
- frecuencia de muestreo y rango dinámico del acelerómetro;
- identificación de los sensores disponibles;
- un informe individual en HTML;
- un Excel individual y un Excel global de control;
- una estructura ordenada por código de participante.


## Dos formas de instalación
