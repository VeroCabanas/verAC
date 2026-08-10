# ========================================================================================================================
# INSTALADOR DE ENTORNO verAC + AUTO-COMPROBACIÓN  (IMPaCT)
#        BY VERONICA CABANAS SANCHEZ
# ------------------------------------------------------------------------------------------------------------------------
# Lo ejecuta UNA VEZ el servicio de informática (o quien tenga permisos), a través de Instalador_verAC.bat.
#   1) Instala los paquetes R imprescindibles desde los binarios .zip locales (sin Internet).
#   2) AUTO-COMPROBACIÓN: lee y analiza los .bin de prueba de scripts/arch_prueba EN MEMORIA
#      (sin generar Excel/HTML, sin mover archivos, sin crear carpetas): solo muestra en consola
#      que el entorno funciona y procesa correctamente.
# (La instalación de R en sí la realiza el .bat antes de llamar a este script.)
# ========================================================================================================================

# ---- Localizar carpetas ----
args <- commandArgs(trailingOnly = FALSE)
script_path <- sub("^--file=", "", grep("^--file=", args, value = TRUE))
script_dir <- if (length(script_path) == 0 || script_path == "") getwd() else dirname(normalizePath(script_path, winslash = "/", mustWork = FALSE))
pkg_dir   <- file.path(script_dir, "paquetesR")
prueba_dir <- file.path(script_dir, "arch_prueba")

message("================================================================================")
message("INSTALADOR DE ENTORNO verAC  +  AUTO-COMPROBACIÓN")
message("================================================================================")
message("R en uso : ", R.version.string)
message("Paquetes : ", pkg_dir)
message("")

# =========================================================================
# 1) INSTALACIÓN DE PAQUETES DESDE .zip LOCALES (por pasadas, robusto)
# =========================================================================
# Se instalan TODOS los .zip de la carpeta por pasadas, resolviendo el orden de
# dependencias automáticamente (sin lista fija, que se rompe si cambian las
# dependencias entre versiones de los paquetes).
nucleo <- c("data.table", "openxlsx", "GGIRread", "zip", "webshot2")

if (!dir.exists(pkg_dir)) {
  message("\u26d4 ERROR: no existe la carpeta de paquetes: ", pkg_dir)
  quit(save = "no", status = 2)
}

zip_files <- list.files(pkg_dir, pattern = "\\.zip$", full.names = TRUE, ignore.case = TRUE)
if (length(zip_files) == 0) {
  message("\u26d4 ERROR: no hay archivos .zip en: ", pkg_dir)
  quit(save = "no", status = 2)
}
zip_map <- list()
for (z in zip_files) {
  nm <- sub("_.*$", "", basename(z))
  if (is.null(zip_map[[nm]])) zip_map[[nm]] <- z
}
pendientes <- names(zip_map)

ya_instalado <- function(pkg) requireNamespace(pkg, quietly = TRUE)
base_pkgs <- rownames(installed.packages(priority = "base"))
pendientes <- setdiff(pendientes, base_pkgs)

message("Instalando ", length(pendientes), " paquete(s) desde binarios locales (sin Internet)...")
message("   (se instalan por pasadas, resolviendo el orden de dependencias automaticamente)")
message("")

max_pasadas <- 12
for (pasada in seq_len(max_pasadas)) {
  quedan <- pendientes[!vapply(pendientes, ya_instalado, logical(1))]
  if (length(quedan) == 0) break
  message("-- Pasada ", pasada, ": quedan ", length(quedan), " por instalar --")
  progreso <- FALSE
  for (pkg in quedan) {
    zp <- zip_map[[pkg]]
    if (is.null(zp)) next
    ok <- tryCatch({
      suppressWarnings(install.packages(zp, repos = NULL, type = "win.binary", quiet = TRUE))
      ya_instalado(pkg)
    }, error = function(e) FALSE)
    if (isTRUE(ok)) { progreso <- TRUE; message("   \u2705 ", pkg) }
  }
  if (!progreso) break
}

quedan_final <- pendientes[!vapply(pendientes, ya_instalado, logical(1))]
nucleo_ko <- nucleo[!vapply(nucleo, ya_instalado, logical(1))]

message("")
if (length(nucleo_ko) > 0) {
  message("\u26d4 INSTALACION INCOMPLETA")
  message("   Paquetes clave no disponibles: ", paste(nucleo_ko, collapse = ", "))
  if (length(quedan_final) > 0) message("   Otros no instalados: ", paste(quedan_final, collapse = ", "))
  message("   Revise que la carpeta 'paquetesR' contiene todos los .zip y reintente.")
  quit(save = "no", status = 2)
}
if (length(quedan_final) > 0) {
  message("\u26a0\ufe0f  Aviso: estos paquetes no se instalaron (puede que no sean imprescindibles): ",
          paste(quedan_final, collapse = ", "))
}
message("\u2705 Paquetes imprescindibles instalados y disponibles")
message("")

# Cargar lo necesario para el auto-test
suppressMessages({
  library(GGIRread)
})

# =========================================================================
# 2) AUTO-COMPROBACIÓN: procesar los .bin de prueba EN MEMORIA
#    (sin escribir nada en disco; solo mostrar resultado en consola)
# =========================================================================
message("================================================================================")
message("AUTO-COMPROBACIÓN: análisis de archivos de prueba (en memoria, sin guardar nada)")
message("================================================================================")

if (!dir.exists(prueba_dir)) {
  message("")
  message("================================================================================")
  message("⚠️ NO SE PUDO COMPLETAR LA AUTO-COMPROBACIÓN")
  message("================================================================================")
  message("R y los paquetes parecen haberse instalado correctamente, PERO no se ha podido")
  message("comprobar que verAC procesa bien los archivos, porque no existe la carpeta de prueba:")
  message("   ", prueba_dir)
  message("")
  message("Por tanto, NO se garantiza la correcta ejecución de verAC.")
  message("Recree la carpeta 'scripts\\arch_prueba' con un .bin de muslo (ACEMI) y uno de")
  message("muñeca (ACEMS), y vuelva a ejecutar el instalador.")
  message("================================================================================")
  quit(save = "no", status = 3)
}

bin_prueba <- list.files(prueba_dir, pattern = "\\.bin$", full.names = TRUE, ignore.case = TRUE)
if (length(bin_prueba) == 0) {
  message("")
  message("================================================================================")
  message("⚠️ NO SE PUDO COMPLETAR LA AUTO-COMPROBACIÓN")
  message("================================================================================")
  message("R y los paquetes parecen haberse instalado correctamente, PERO no se ha podido")
  message("comprobar que verAC procesa bien los archivos, porque no hay ningún .bin de prueba en:")
  message("   ", prueba_dir)
  message("")
  message("Por tanto, NO se garantiza la correcta ejecución de verAC.")
  message("Coloque un .bin de muslo (ACEMI) y uno de muñeca (ACEMS) en esa carpeta y vuelva")
  message("a ejecutar el instalador.")
  message("================================================================================")
  quit(save = "no", status = 3)
}

# Detección ligera de posición por nombre (igual criterio que el procesador)
detect_pos <- function(fn) {
  if (grepl("ACEMI", fn, ignore.case = TRUE)) return("muslo")
  if (grepl("ACEMS", fn, ignore.case = TRUE)) return("muñeca")
  return("desconocida")
}

todo_ok <- TRUE
for (bf in bin_prueba) {
  nm <- basename(bf)
  pos <- detect_pos(nm)
  message("── Archivo: ", nm, "  (posición: ", pos, ")")
  res <- tryCatch(
    GGIRread::readParmayMatrix(bf, read_gyro = TRUE, read_heart = TRUE),
    error = function(e) { message("   ❌ Error de lectura: ", conditionMessage(e)); NULL }
  )
  if (is.null(res) || is.null(res$data) || !all(c("acc_x","acc_y","acc_z") %in% names(res$data))) {
    message("   ❌ No se pudo leer/decodificar correctamente (estructura inválida).")
    todo_ok <- FALSE; next
  }
  d <- res$data
  if (length(d$acc_x) == 0) {
    message("   ❌ El archivo se leyó pero no contiene datos de acelerometría (0 muestras).")
    todo_ok <- FALSE; next
  }
  sf <- tryCatch(res$header$sf, error = function(e) 25)
  dyn <- tryCatch(res$header$acc_dynrange, error = function(e) NA)
  nmuestras <- length(d$acc_x)
  tini <- tryCatch(format(as.POSIXct(min(d$time), origin="1970-01-01"), "%d/%m/%Y %H:%M:%S"), error=function(e) "N/D")
  message("   ✅ Lectura correcta | muestras acc: ", nmuestras, " | ", sf, " Hz | rango ", dyn, " g | inicio ", tini)
}

message("")
message("================================================================================")
if (todo_ok) {
  message("✅ AUTO-COMPROBACIÓN SUPERADA: el entorno está instalado y procesa correctamente.")
  message("   Ya puede usarse verAC con normalidad (verAC_v3.10.bat) para procesar archivos.")
  message("================================================================================")
  quit(save = "no", status = 0)
} else {
  message("⛔ LA AUTO-COMPROBACIÓN ENCONTRÓ ERRORES")
  message("================================================================================")
  message("R y los paquetes parecen instalados, pero al menos un archivo de prueba NO se pudo")
  message("leer o procesar correctamente (ver detalles arriba).")
  message("Esto indica que algo no funciona bien y NO se garantiza la correcta ejecución de verAC.")
  message("")
  message("Qué hacer:")
  message("   1. Vuelva a ejecutar el instalador (Instalador_verAC.bat).")
  message("   2. Si el problema persiste, contacte con la desarrolladora del programa:")
  message("      veronica.cabanas@uam.es")
  message("================================================================================")
  quit(save = "no", status = 4)
}
