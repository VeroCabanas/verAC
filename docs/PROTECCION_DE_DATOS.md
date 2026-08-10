# Protección de datos y publicación abierta

Los archivos de acelerometría y los resultados derivados pueden contener códigos de participante, fechas de registro y otros metadatos sujetos al protocolo del estudio.

## No publicar

- archivos `.bin` de participantes;
- Excel o informes HTML/PDF generados con datos reales;
- carpetas `archivos bin procesados`;
- capturas de pantalla con códigos o fechas identificables;
- rutas locales o registros de ejecución que revelen información personal.

El `.gitignore` del repositorio excluye estas extensiones y carpetas, pero esa protección no sustituye una revisión manual antes de cada publicación.

## Archivos de prueba

Los dos `.BIN` de `scripts/arch_prueba` fueron creados expresamente para validar el programa, no proceden de participantes y están autorizados para su difusión. Son las únicas excepciones a la exclusión general de archivos `.bin` definida en `.gitignore`.

## Incidencias

No adjunte datos reales a incidencias públicas. Utilice un ejemplo sintético o describa el problema sin incluir identificadores.
