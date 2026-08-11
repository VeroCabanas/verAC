required_r <- "4.4.0"
required <- c(
  data.table = "1.18.4",
  openxlsx = "4.2.8.1",
  GGIRread = "1.0.8",
  zip = "3.0.0"
)

if (getRversion() < required_r) {
  stop(paste("verAC requiere R", required_r, "o posterior."), call. = FALSE)
}

repos <- c(CRAN = "https://cloud.r-project.org")
if (!requireNamespace("remotes", quietly = TRUE)) {
  install.packages("remotes", repos = repos)
}

for (package in names(required)) {
  target <- required[[package]]
  installed <- if (requireNamespace(package, quietly = TRUE)) {
    as.character(utils::packageVersion(package))
  } else {
    NA_character_
  }
  if (is.na(installed) || installed != target) {
    message("Instalando ", package, " ", target, "...")
    remotes::install_version(
      package,
      version = target,
      repos = repos,
      upgrade = "never",
      dependencies = NA
    )
  }
}

actual <- vapply(names(required), function(package) {
  as.character(utils::packageVersion(package))
}, character(1))

if (!all(actual == required)) {
  stop(
    paste(
      "No se obtuvieron todas las versiones previstas:",
      paste(names(actual), actual, sep = "=", collapse = "; ")
    ),
    call. = FALSE
  )
}

message(
  "Dependencias instaladas: ",
  paste(names(actual), actual, sep = " ", collapse = "; ")
)
