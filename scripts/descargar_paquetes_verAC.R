# =============================================================================
#  descargar_paquetes_verAC.R
# -----------------------------------------------------------------------------
#  Descarga TODOS los paquetes de R que necesita verAC (con sus dependencias)
#  como archivos .zip (binarios de Windows), a una carpeta local, para instalar
#  despues SIN conexion a internet en los centros.
#
#  CÓMO USARLO (una sola vez, en un equipo CON internet):
#   1. Este equipo debe tener el MISMO R que usa verAC (R 4.6.1).
#   2. Ejecutarlo con ese R. Por ejemplo, desde cmd:
#        "E:\IMPaCT_2026\verAC\scripts\R\R-4.6.1\bin\x64\Rscript.exe" ^
#           "E:\IMPaCT_2026\verAC\scripts\descargar_paquetes_verAC.R"
#   3. Al terminar, la carpeta 'paquetesR' contendra todos los .zip necesarios.
# =============================================================================

args <- commandArgs(trailingOnly = FALSE)
script_path <- sub("^--file=", "", args[grep("^--file=", args)])
if (length(script_path) == 0) {
  script_dir <- getwd()
} else {
  script_dir <- dirname(normalizePath(script_path))
}
dest_dir <- file.path(script_dir, "paquetesR")
if (!dir.exists(dest_dir)) dir.create(dest_dir, recursive = TRUE, showWarnings = FALSE)

repo <- "https://cloud.r-project.org"
options(repos = c(CRAN = repo))

cat("============================================================\n")
cat("  Descarga de paquetes para verAC (instalacion offline)\n")
cat("============================================================\n")
cat("Version de R en uso: ", R.version.string, "\n", sep = "")
cat("Carpeta de destino:  ", dest_dir, "\n\n", sep = "")

if (getRversion() < "4.6.0" || getRversion() >= "4.7.0") {
  cat("ATENCION: este R no es 4.6.x. Los .zip descargados podrian no ser\n")
  cat("    compatibles con el R 4.6.1 de verAC. Ejecutelo con el R de verAC.\n\n")
}

# --- Paquetes de primer nivel que usa verAC ---
#  data.table, openxlsx, GGIRread: nucleo de lectura/escritura y QC.
#  webshot2: generacion de PDF (usa Chrome/Chromium).
paquetes_top <- c(
  "data.table",
  "openxlsx",
  "GGIRread",
  "webshot2"
)

cat("Resolviendo el arbol completo de dependencias...\n")
ap <- available.packages(repos = repo, type = "win.binary")
deps <- tools::package_dependencies(
  packages = paquetes_top,
  db = ap,
  which = c("Depends", "Imports", "LinkingTo"),
  recursive = TRUE
)
todos <- unique(c(paquetes_top, unlist(deps, use.names = FALSE)))

base_pkgs <- rownames(installed.packages(priority = "base"))
todos <- setdiff(todos, base_pkgs)
todos <- sort(todos)

cat("Paquetes a descargar (", length(todos), "):\n", sep = "")
cat(paste(" -", todos), sep = "\n")
cat("\n\nDescargando binarios de Windows (.zip) a la carpeta de destino...\n")

ok <- tryCatch({
  download.packages(pkgs = todos, destdir = dest_dir, repos = repo, type = "win.binary")
}, error = function(e) {
  cat("\nError durante la descarga: ", conditionMessage(e), "\n", sep = "")
  NULL
})

cat("\n============================================================\n")
if (!is.null(ok)) {
  zips <- list.files(dest_dir, pattern = "\\.zip$", full.names = FALSE)
  cat("Descarga finalizada.\n")
  cat("   Paquetes solicitados: ", length(todos), "\n", sep = "")
  cat("   Archivos .zip en la carpeta: ", length(zips), "\n", sep = "")
  cat("   Carpeta: ", dest_dir, "\n", sep = "")
  descargados <- sub("_.*$", "", zips)
  faltan <- setdiff(todos, descargados)
  if (length(faltan) > 0) {
    cat("\nNo se descargaron estos paquetes (revisar):\n")
    cat(paste(" -", faltan), sep = "\n")
    cat("\n")
  } else {
    cat("\n   Todos los paquetes se descargaron correctamente.\n")
  }
} else {
  cat("La descarga no se completo. Revise la conexion a internet y reintentelo.\n")
}
cat("============================================================\n")
