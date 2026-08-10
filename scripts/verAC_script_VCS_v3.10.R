# ========================================================================================================================
# COMPROBACION INICIAL DE ARCHIVOS ORIGINALES IMPACT: PRE-PROCESADO PARA EXTRAER DATOS BÁSICOS DE ARCHIVOS .BIN (MATRIX) #
# EJ.: ARCHIVO CORRUPTO, FECHAS DE REGISTROS, DURACIÓN DEL REGISTRO, ETC. DE ARCHIVOS DE ACELEROMETRÍA (MUÑECA O MUSLO)  #
#                            versión 3 (6/2026)      BY VERÓNICA CABANAS-SÁNCHEZ                                              #
# ========================================================================================================================


# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y PARÁMETROS
# -----------------------------------------------------------------------------

# Configuración principal
CONFIG <- list(
  guardar_csv = FALSE,
  sobreescribir_csv = TRUE,
  csv_positions = c("muslo"),  # Opciones: "muslo", "muñeca", "desconocida"
  guardar_info_paquetes = FALSE,   # Detalle de paquetes en los Excel (solo para revisión de la investigadora principal).
                                   # TRUE  = añade al final de los Excel (individual y global) las columnas Paquetes_declarados, Paquetes_observados, Diferencia_paquetes y Paquetes_imputados.
                                   # FALSE = uso habitual de los técnicos: no aparecen esas columnas.
  timezone = "Europe/Madrid",
  max_espera_excel = 600,  # seg
  intervalo_espera = 30,    # seg
  generar_pdf = FALSE, # poner FALSE si no se quieren convertir los informes html a pdf

  # Opciones de reprocesamiento
  reprocesar_existentes = TRUE,  # TRUE = reprocesa archivos ya procesados, FALSE = los salta

  # Opciones de limpieza de duplicados en Excels
  limpiar_duplicados_excel_global = FALSE,     # TRUE = elimina filas anteriores del mismo archivo en Excel global, FALSE = mantiene historial
  limpiar_duplicados_excel_individual = TRUE, # TRUE = elimina filas anteriores del mismo archivo en Excel individuales, FALSE = mantiene historial

  # Opciones de backup específicas por tipo de archivo
  backup_archivos_bin = FALSE,        # TRUE = crea backups de archivos .bin (con timestamps), FALSE = sobreescribe el archivo .bin en la carpeta del participante cuando se reprocesa el archivo
  backup_archivos_html = FALSE,       # TRUE = crea backups de archivos .html (con timestamps), FALSE = sobreescribe el informe .html en la carpeta del participante cuando se reprocesa el archivo
  backup_excel_individual = FALSE,    # TRUE = crea backups de Excel individuales de cada participante (con timestamps); FALSE = sobreescribe el Excel individual de cada participante cuando se reprocesa el archivo (elimina/mantiene historial según parámetro limpiar_duplicados_excel_global)
  backup_excel_global = FALSE         # TRUE = crea backups del Excel global (con timestamps); FALSE = sobreescribe el Excel global cuando se reprocesa el archivo (elimina/mantiene historial según parámetro limpiar_duplicados_excel_individual)
)

#' Eliminar relaciones internas colgantes generadas por openxlsx 4.2.8.1
#'
#' Algunas versiones de openxlsx pueden escribir referencias a componentes de
#' impresión o dibujo que no se incorporan al .xlsx. Excel suele repararlas de
#' forma silenciosa, pero los validadores Open XML estrictos rechazan el libro.
#' Esta función elimina exclusivamente relaciones cuyo destino no existe y sus
#' referencias, sin modificar los datos ni el formato de las celdas.
repair_openxlsx_package <- function(xlsx_path) {
  if (!file.exists(xlsx_path)) return(invisible(FALSE))
  if (!requireNamespace("zip", quietly = TRUE)) {
    warning("No se pudo validar la estructura interna del Excel: falta el paquete 'zip'.")
    return(invisible(FALSE))
  }

  stage <- tempfile("verac_xlsx_")
  dir.create(stage, recursive = TRUE)
  on.exit(unlink(stage, recursive = TRUE, force = TRUE), add = TRUE)
  utils::unzip(xlsx_path, exdir = stage)

  rel_files <- list.files(stage, pattern = "\\.rels$", recursive = TRUE,
                          full.names = TRUE, all.files = TRUE)
  rel_files <- rel_files[file.exists(rel_files) & !dir.exists(rel_files)]
  for (rel_file in rel_files) {
    rel_txt <- paste(readLines(rel_file, warn = FALSE, encoding = "UTF-8"), collapse = "")
    rel_nodes <- regmatches(rel_txt, gregexpr("<Relationship\\b[^>]*/>", rel_txt, perl = TRUE))[[1]]
    if (length(rel_nodes) == 0 || identical(rel_nodes, "")) next

    rel_dir <- dirname(rel_file)
    source_dir <- dirname(rel_dir)
    source_name <- sub("\\.rels$", "", basename(rel_file))
    source_xml <- file.path(source_dir, source_name)
    source_txt <- if (file.exists(source_xml) && !dir.exists(source_xml)) {
      paste(readLines(source_xml, warn = FALSE, encoding = "UTF-8"), collapse = "")
    } else NULL

    for (node in rel_nodes) {
      if (grepl('TargetMode="External"', node, fixed = TRUE)) next
      target <- sub('.*Target="([^"]+)".*', "\\1", node)
      rel_id <- sub('.*Id="([^"]+)".*', "\\1", node)
      target_path <- normalizePath(file.path(source_dir, target), winslash = "/",
                                   mustWork = FALSE)
      if (!file.exists(target_path)) {
        rel_txt <- sub(node, "", rel_txt, fixed = TRUE)
        if (!is.null(source_txt) && nzchar(rel_id)) {
          draw_pattern <- paste0('<(?:drawing|legacyDrawing)\\b[^>]*r:id="',
                                 rel_id, '"[^>]*/>')
          source_txt <- gsub(draw_pattern, "", source_txt, perl = TRUE)
          source_txt <- gsub(paste0('\\s+r:id="', rel_id, '"'), "", source_txt,
                             perl = TRUE)
        }
      }
    }
    writeLines(rel_txt, rel_file, useBytes = TRUE)
    if (!is.null(source_txt)) writeLines(source_txt, source_xml, useBytes = TRUE)
  }

  content_types <- file.path(stage, "[Content_Types].xml")
  if (file.exists(content_types)) {
    ct <- paste(readLines(content_types, warn = FALSE, encoding = "UTF-8"), collapse = "")
    overrides <- regmatches(ct, gregexpr("<Override\\b[^>]*/>", ct, perl = TRUE))[[1]]
    if (length(overrides) > 0 && !identical(overrides, "")) {
      for (node in overrides) {
        part <- sub('.*PartName="/([^"]+)".*', "\\1", node)
        if (!file.exists(file.path(stage, part))) ct <- sub(node, "", ct, fixed = TRUE)
      }
    }
    writeLines(ct, content_types, useBytes = TRUE)
  }

  rebuilt <- tempfile(fileext = ".xlsx")
  on.exit(unlink(rebuilt, force = TRUE), add = TRUE)
  files <- list.files(stage, recursive = TRUE, all.files = TRUE, no.. = TRUE)
  zip::zipr(rebuilt, files = files, root = stage, mode = "mirror")
  ok <- file.copy(rebuilt, xlsx_path, overwrite = TRUE)
  if (!ok) stop("No se pudo guardar el Excel validado: ", xlsx_path)
  invisible(TRUE)
}

# -----------------------------------------------------------------------------
# Refuerzo de la librería local (blindaje)
# Antepone la librería del R portable del propio programa (scripts/R/R-x.y.z/library)
# a .libPaths(), para garantizar que se usan los paquetes instalados ahí y no una
# librería de usuario que pudiera tener versiones distintas.
# -----------------------------------------------------------------------------
local({
  args <- commandArgs(trailingOnly = FALSE)
  sp <- sub("^--file=", "", grep("^--file=", args, value = TRUE))
  sd <- if (length(sp) == 0 || sp == "") getwd() else dirname(normalizePath(sp, winslash = "/", mustWork = FALSE))
  base_dir <- if (basename(sd) == "scripts") dirname(sd) else sd
  r_root <- file.path(base_dir, "scripts", "R")
  if (dir.exists(r_root)) {
    libs <- list.dirs(r_root, recursive = TRUE, full.names = TRUE)
    libs <- libs[basename(libs) == "library"]
    if (length(libs) > 0) {
      lib_local <- normalizePath(libs[1], winslash = "/", mustWork = FALSE)
      .libPaths(c(lib_local, .libPaths()))
    }
  }
})

# -----------------------------------------------------------------------------
# 2. FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------

#' Extraer código de participante del nombre del archivo
get_participant_code <- function(filename) {
  # Buscar desde el primer guión hasta el último punto
  # Ejemplo: MD ACEMS1-13010532.bin -> 13010532
  if (grepl("-", filename) && grepl("\\.", filename)) {
    # Extraer desde el primer guión hasta el último punto
    pattern <- ".*?-([^.]+)\\..*"
    if (grepl(pattern, filename)) {
      return(sub(pattern, "\\1", filename))
    }
  }
  return("desconocido")
}

#' Verificar si un archivo ya fue procesado
is_file_already_processed <- function(filename, participant_dir) {
  if (!dir.exists(participant_dir)) return(FALSE)

  # Buscar si el archivo .bin específico ya existe en la carpeta del participante
  bin_path <- file.path(participant_dir, filename)
  return(file.exists(bin_path))
}

#' Crear backup de un archivo si está configurado según el tipo
create_backup_if_needed <- function(filepath, file_type) {
  if (!file.exists(filepath)) return()

  # Verificar si el backup está habilitado para este tipo de archivo
  backup_enabled <- switch(file_type,
                           "bin" = CONFIG$backup_archivos_bin,
                           "html" = CONFIG$backup_archivos_html,
                           "excel_individual" = CONFIG$backup_excel_individual,
                           "excel_global" = CONFIG$backup_excel_global,
                           FALSE  # Por defecto no hacer backup si no se reconoce el tipo
  )

  if (!backup_enabled) return()

  # Generar timestamp en formato YYYYMMDD_HHMMSS
  timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")

  # Extraer nombre base y extensión
  file_dir <- dirname(filepath)
  file_name <- basename(filepath)

  # Separar nombre y extensión
  if (grepl("\\.", file_name)) {
    name_parts <- strsplit(file_name, "\\.")[[1]]
    if (length(name_parts) >= 2) {
      base_name <- paste(name_parts[1:(length(name_parts)-1)], collapse = ".")
      extension <- name_parts[length(name_parts)]
      backup_name <- paste0(base_name, "_", timestamp, "_bak.", extension)
    } else {
      # Archivo sin extensión (raro, pero por seguridad)
      backup_name <- paste0(file_name, "_", timestamp, "_bak")
    }
  } else {
    # Archivo sin extensión
    backup_name <- paste0(file_name, "_", timestamp, "_bak")
  }

  backup_path <- file.path(file_dir, backup_name)

  tryCatch({
    file.copy(filepath, backup_path, overwrite = TRUE)
    message("🗄️ Backup creado: ", backup_name)
  }, error = function(e) {
    message("⚠️ No se pudo crear backup de ", basename(filepath), ": ", e$message)
  })
}

#' Limpiar duplicados del Excel si está configurado
clean_excel_duplicates <- function(excel_path, device_id, clean_duplicates) {
  if (!clean_duplicates || !file.exists(excel_path)) return()

  tryCatch({
    # Leer datos existentes
    existing_data <- openxlsx::read.xlsx(excel_path, sheet = 1)

    if (nrow(existing_data) > 0) {
      # Filtrar filas que NO correspondan al archivo actual
      filtered_data <- existing_data[existing_data$ID_archivo != device_id, ]

      if (nrow(filtered_data) < nrow(existing_data)) {
        message("🧹 Eliminando datos anteriores del archivo ", device_id, " del Excel")

        # Crear nuevo workbook con datos filtrados
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, "Sheet1")

        if (nrow(filtered_data) > 0) {
          openxlsx::writeData(wb, sheet = 1, x = filtered_data, startRow = 1, colNames = TRUE)
        } else {
          # Si no quedan datos, solo escribir encabezados
          headers <- data.frame(
            ID_archivo = character(0),
            Fecha_procesamiento = character(0),
            Posición = character(0),
            ESTADO_ARCHIVO = character(0),
            Hora_inicio_grabación = as.POSIXct(character(0)),
            Hora_fin_grabación = as.POSIXct(character(0)),
            Duración_grabación_min = numeric(0),
            Duración_grabación_días = numeric(0),
            Acelerómetro = character(0),
            Rango_dinámico_acc = numeric(0),
            Frecuencia_muestreo = numeric(0),
            Giróscopo = character(0),
            Frecuencia_cardiaca = character(0),
            Temperatura_corporal = character(0),
            Temperatura_ambiente = character(0)
          )
          if (isTRUE(CONFIG$guardar_info_paquetes)) {
            headers$Paquetes_declarados <- integer(0)
            headers$Paquetes_observados <- integer(0)
            headers$Diferencia_paquetes <- integer(0)
            headers$Paquetes_imputados  <- integer(0)
          }
          openxlsx::writeData(wb, sheet = 1, x = headers, startRow = 1, colNames = TRUE)
        }

        openxlsx::saveWorkbook(wb, excel_path, overwrite = TRUE)
        repair_openxlsx_package(excel_path)
      }
    }
  }, error = function(e) {
    message("⚠️ Error limpiando duplicados del Excel: ", e$message)
  })
}

#' Crear directorios necesarios (nueva estructura)
setup_directories <- function() {
  # Detectar si estamos en la carpeta 'scripts' y ajustar la ruta base
  base_dir <- getwd()
  if (basename(base_dir) == "scripts") {
    base_dir <- dirname(base_dir)  # Subir un nivel
    message("📂 Detectada ejecución desde carpeta 'scripts', ajustando rutas al directorio padre")
  }

  dirs <- list(
    bin = file.path(base_dir, "archivos bin"),
    processed_base = file.path(base_dir, "archivos bin procesados"),
    csv = if(CONFIG$guardar_csv) file.path(base_dir, "archivos csv") else NULL,
    excel_results = file.path(base_dir, "Excel de resultados del procesado")
  )

  # Crear directorios base que no existen
  for (dir_name in names(dirs)) {
    if (!is.null(dirs[[dir_name]]) && !dir.exists(dirs[[dir_name]])) {
      dir.create(dirs[[dir_name]], recursive = TRUE)
      message("📁 Directorio creado: ", basename(dirs[[dir_name]]))
    }
  }

  return(dirs)
}

#' Crear/configurar directorio específico del participante
setup_participant_directory <- function(base_dirs, participant_code) {
  participant_dir <- file.path(base_dirs$processed_base, participant_code)
  if (!dir.exists(participant_dir)) {
    dir.create(participant_dir, recursive = TRUE)
    message("📁 Directorio de participante creado: ", participant_code)
  }
  return(participant_dir)
}

#' Determinar posición del dispositivo basado en el nombre del archivo
get_device_position <- function(filename) {
  filename_upper <- toupper(filename)
  if (grepl("ACEMI", filename_upper)) {
    return("muslo")
  } else if (grepl("ACEMS", filename_upper)) {
    return("muñeca")
  } else {
    return("desconocida")
  }
}

#' Verificar si un archivo está en uso (con supresión de avisos)
is_file_in_use <- function(filepath) {
  if (!file.exists(filepath)) return(FALSE)

  tryCatch({
    suppressWarnings({
      con <- file(filepath, open = "r+b")
      if (isOpen(con)) {
        close(con)
        return(FALSE)  # No está en uso
      }
      return(TRUE)  # Está en uso
    })
  }, error = function(e) {
    return(TRUE)
  })
}

#' Esperar hasta que el archivo EXCEL esté disponible
wait_for_excel <- function(filepath, max_wait = CONFIG$max_espera_excel) {
  if (!file.exists(filepath)) return(TRUE)

  # Verificación rápida inicial - si no está en uso, continuar
  if (!is_file_in_use(filepath)) return(TRUE)

  # Mostrar mensaje si  está en uso
  start_time <- Sys.time()

  while (is_file_in_use(filepath)) {
    elapsed <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))

    if (elapsed > max_wait) {
      stop("⛔ [TIEMPO_AGOTADO_EXCEL] Tiempo de espera agotado. Cierre el archivo Excel y reinicie el proceso.")
    }

    message("⚠️ El archivo de Excel está abierto. Por favor, ciérrelo para continuar. Esperando... (", round(elapsed), " segundos transcurridos)")
    Sys.sleep(CONFIG$intervalo_espera)
  }

  return(TRUE)
}

#' Esperar hasta que el archivo HTML esté disponible
wait_for_html <- function(filepath, max_wait = CONFIG$max_espera_excel) {
  if (!file.exists(filepath)) return(TRUE)

  # Verificación rápida inicial - si no está en uso, continuar
  if (!is_file_in_use(filepath)) return(TRUE)

  # Mostrar mensaje si  está en uso
  message("⚠️ Archivo HTML en uso, esperando que se cierre...")
  start_time <- Sys.time()

  while (is_file_in_use(filepath)) {
    elapsed <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))

    if (elapsed > max_wait) {
      stop("⛔ [TIEMPO_AGOTADO_HTML] Tiempo de espera agotado. Cierre el archivo HTML y reinicie el proceso.")
    }

    message("⚠️ El archivo HTML está abierto. Por favor, ciérrelo para continuar. Esperando... (", round(elapsed), " segundos transcurridos)")
    Sys.sleep(CONFIG$intervalo_espera)
  }

  return(TRUE)
}

#' Esperar hasta que el archivo PDF esté disponible
wait_for_pdf <- function(filepath, max_wait = CONFIG$max_espera_excel) {
  if (!file.exists(filepath)) return(TRUE)

  # Verificación rápida inicial - si no está en uso, continuar
  if (!is_file_in_use(filepath)) return(TRUE)

  # Solo mostrar mensaje inicial si realmente está en uso
  message("⚠️ Archivo PDF en uso, esperando que se cierre...")
  start_time <- Sys.time()

  while (is_file_in_use(filepath)) {
    elapsed <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))

    if (elapsed > max_wait) {
      stop("⛔ [TIEMPO_AGOTADO_PDF] Tiempo de espera agotado. Cierre el archivo PDF y reinicie el proceso.")
    }

    message("⚠️ El archivo PDF está abierto. Por favor, ciérrelo para continuar. Esperando... (", round(elapsed), " segundos transcurridos)")
    Sys.sleep(CONFIG$intervalo_espera)
  }

  return(TRUE)
}

#' Añade (si CONFIG$guardar_info_paquetes es TRUE) las cuatro columnas de detalle de
#' paquetes al final de un data.table de resultados. Si es FALSE, devuelve el dt sin cambios.
#' Pensado para uso de la investigadora principal en la revisión global de calidad.
add_info_paquetes <- function(dt, rdata) {
  if (!isTRUE(CONFIG$guardar_info_paquetes)) return(dt)
  declarados <- if (is.null(rdata$n_declarados)) NA_integer_ else rdata$n_declarados
  observados <- if (is.null(rdata$n_observados)) NA_integer_ else rdata$n_observados
  diferencia <- if (!is.na(declarados) && !is.na(observados)) observados - declarados else NA_integer_
  imputados  <- if (is.null(rdata$n_imputados)) NA_integer_ else rdata$n_imputados
  dt$Paquetes_declarados <- declarados
  dt$Paquetes_observados <- observados
  dt$Diferencia_paquetes <- diferencia
  dt$Paquetes_imputados  <- imputados
  dt
}

#' Etiqueta visible del estado del archivo (para Excel e informes).
#' Mantiene el valor interno ("válido"/"semi-válido"/"corrupto") intacto en la lógica,
#' y solo cambia cómo se MUESTRA: "semi-válido" -> "(semi) válido".
estado_label <- function(estado) {
  if (is.null(estado) || is.na(estado)) return(estado)
  switch(as.character(estado),
         "semi-válido" = "(semi) válido",
         "válido"      = "válido",
         "corrupto"    = "corrupto",
         as.character(estado))
}

#' Transformación de ejes Matrix -> convención ActiPASS/Axivity (muslo, asumiendo upright).
#' Se aplica SIEMPRE a los CSV de muslo. ActiPASS detecta y corrige automáticamente la
#' orientación (upright/inverted) durante su análisis.
#'   X_ActiPASS = -Y_Matrix (vertical, + hacia el pie)
#'   Y_ActiPASS =  X_Matrix (lateral)
#'   Z_ActiPASS =  Z_Matrix (hacia la piel)
#' @return data.frame(x, y, z), o NULL si no hay datos de acelerometría
rotate_axes_basic <- function(bin_data) {
  if (is.null(bin_data) || is.null(bin_data$data)) return(NULL)
  d <- bin_data$data
  if (!all(c("acc_x","acc_y","acc_z") %in% names(d))) return(NULL)
  data.frame(
    x = d$acc_y * (-1),  # X_ActiPASS = -Y_Matrix
    y = d$acc_x,         # Y_ActiPASS =  X_Matrix
    z = d$acc_z          # Z_ActiPASS =  Z_Matrix
  )
}

#' Procesar archivo BIN individual


process_bin_file <- function(filepath) {
  device_id <- basename(filepath)
  position <- get_device_position(device_id)

  # ------------------------------------------------------------------------
  # LECTURA DEL ARCHIVO con distinción robusta de tres situaciones:
  #   (a) ARCHIVO CORRUPTO  : GGIRread no puede decodificarlo porque los datos
  #       están dañados. Da un error ESPECÍFICO y reconocible (p.ej. cabecera
  #       inválida). -> se marca "corrupto" y se genera informe normal.
  #   (b) PROBLEMA DE HERRAMIENTA/ENTORNO : GGIRread no funciona (dependencia
  #       bloqueada, namespace, función no disponible, etc.). El error NO es de
  #       los específicos de datos. -> NO se marca corrupto; NO se genera informe;
  #       se devuelve un marcador para avisar de "error inesperado, revisar instalación".
  #
  # Estrategia (lista blanca): SOLO los errores compatibles con datos dañados
  # marcan corrupto. Cualquier otro error se considera problema de herramienta.
  # Esto es lo seguro: ante un error no reconocido, no asumimos que el archivo
  # esté dañado (evita marcar como corruptos archivos sanos por fallos de entorno).
  # ------------------------------------------------------------------------
  tool_error_msg <- NULL  # si se rellena, hubo un problema de herramienta/entorno (no de dato)

  bin_data <- tryCatch({
    GGIRread::readParmayMatrix(filepath, read_gyro = TRUE, read_heart = TRUE)
  }, error = function(e) {
    msg <- conditionMessage(e)
    # Patrones de error que SÍ indican un archivo de datos corrupto/dañado
    # (mensajes propios de la decodificación de GGIRread al leer un .bin Matrix dañado).
    corrupt_patterns <- paste(
      "Invalid header recognition string",   # cabecera no es 'MDTC' (corrupto típico)
      "header",                               # otros problemas de cabecera
      "checksum",                             # fallo de checksum de datos
      "recognition string",
      "unexpected end", "unexpected EOF", "end of file", "fin de archivo",
      "subscript out of bounds", "índice fuera", "out of bounds",
      "invalid", "corrupt", "malformed", "cannot read", "no se puede leer",
      sep = "|"
    )
    if (grepl(corrupt_patterns, msg, ignore.case = TRUE)) {
      # Error compatible con DATOS DAÑADOS -> archivo corrupto. Devolver NULL.
      return(NULL)
    }
    # Cualquier OTRO error -> PROBLEMA DE HERRAMIENTA/ENTORNO (no es corrupción del archivo).
    # No marcamos corrupto: registramos el mensaje para avisar y no generar informe.
    tool_error_msg <<- msg
    return(NULL)
  })

  # ------------------------------------------------------------------------
  # Si hubo PROBLEMA DE HERRAMIENTA/ENTORNO: no es un archivo corrupto. No se
  # genera informe ni se incluye en Excel; se devuelve un estado especial para
  # que el flujo principal avise y, si procede, detenga el procesamiento.
  # ------------------------------------------------------------------------
  if (!is.null(tool_error_msg)) {
    message("")
    message("⛔ ERROR INESPERADO al leer el archivo: ", device_id)
    message("   El problema NO parece ser que el archivo esté corrupto, sino un fallo de")
    message("   la herramienta o del entorno (por ejemplo, un paquete o dependencia de R")
    message("   bloqueado o no disponible). Detalle técnico: ", tool_error_msg)
    message("   ACCIÓN: revise la instalación del entorno (ejecute 'Instalador_verAC.bat').")
    message("   No se genera informe para este archivo para no marcarlo como corrupto por error.")
    return(list(result = NULL, data = NULL, tool_error = tool_error_msg))
  }

  # ------------------------------------------------------------------------
  # Validar la ESTRUCTURA del resultado (si GGIRread devolvió objeto).
  # IMPORTANTE: un archivo SIN acelerometría NO es corrupto: el Matrix puede
  # programarse para registrar solo HR y/o temperatura (error de protocolo, pero
  # el archivo tiene cientos de miles de registros válidos). La ausencia de acc se
  # avisa aparte (mensaje específico en el informe). Por tanto, aquí solo marcamos
  # corrupto si NO hay NINGÚN dato de NINGÚN indicador (sin $data o con 0 filas).
  # ------------------------------------------------------------------------
  estructura_ok <- TRUE
  if (!is.null(bin_data)) {
    if (is.null(bin_data$data) ||
        !is.data.frame(bin_data$data) ||
        nrow(bin_data$data) == 0) {
      estructura_ok <- FALSE
    }
  }

  # ------------------------------------------------------------------------
  # Determinar ESTADO del archivo:
  #   - "corrupto"     : bin_data es NULL o estructura inválida (no hay datos).
  #   - "semi-válido"  : se leyó bien PERO presenta alguno de estos problemas de calidad:
  #         (a) TRUNCAMIENTO: declared_packets != observed_packets (faltan paquetes
  #             enteros; el archivo tiene menos de los que su cabecera declara, típico
  #             de una transferencia interrumpida).
  #         (b) IMPUTACIÓN: algún paquete presente falló el checksum (datos dañados que
  #             GGIRread rellenó con valores artificiales). Se detecta por imputed==TRUE
  #             o checksum_pass==FALSE en el QClog.
  #   - "válido"       : se leyó bien, sin truncamiento ni paquetes imputados.
  # ------------------------------------------------------------------------
  estado_archivo <- "válido"
  n_declarados <- NA_integer_  # total de paquetes declarados en la cabecera
  n_observados <- NA_integer_  # total de paquetes realmente observados
  n_faltantes <- 0L     # paquetes que FALTAN (observed < declared): truncamiento
  n_sobrantes <- 0L     # paquetes de MÁS (observed > declared): cabecera mal grabada
  n_imputados <- 0L     # paquetes presentes pero dañados (checksum fallido / imputados)
  if (is.null(bin_data) || !estructura_ok) {
    estado_archivo <- "corrupto"
  } else {
    declared <- tryCatch(bin_data$QClog$declared_packets, error = function(e) NULL)
    observed <- tryCatch(bin_data$QClog$observed_packets, error = function(e) NULL)
    imputed  <- tryCatch(bin_data$QClog$imputed,          error = function(e) NULL)
    chec_ok  <- tryCatch(bin_data$QClog$checksum_pass,    error = function(e) NULL)

    # (a) Discrepancia declarados vs observados.
    #     observed < declared -> FALTAN paquetes (truncamiento, p.ej. transferencia cortada).
    #     observed > declared -> SOBRAN paquetes (la cabecera del dispositivo grabó mal el
    #                            número de paquetes; el archivo tiene más datos de los declarados,
    #                            pero son válidos y analizables).
    hay_discrepancia <- FALSE
    if (!is.null(declared) && !is.null(observed) && length(declared) > 0 && length(observed) > 0) {
      d_max <- max(declared, na.rm = TRUE)
      o_max <- max(observed, na.rm = TRUE)
      if (!is.na(d_max)) n_declarados <- as.integer(d_max)
      if (!is.na(o_max)) n_observados <- as.integer(o_max)
      if (!is.na(d_max) && !is.na(o_max) && d_max != o_max) {
        hay_discrepancia <- TRUE
        if (o_max < d_max) n_faltantes <- d_max - o_max   # faltan
        if (o_max > d_max) n_sobrantes <- o_max - d_max   # sobran
      }
    }

    # (b) Imputación: paquetes presentes con checksum fallido (datos dañados imputados)
    hay_imputacion <- FALSE
    if (!is.null(imputed) && length(imputed) > 0) {
      n_imputados <- sum(imputed %in% TRUE, na.rm = TRUE)
    }
    if (n_imputados == 0 && !is.null(chec_ok) && length(chec_ok) > 0) {
      n_imputados <- sum(chec_ok %in% FALSE, na.rm = TRUE)
    }
    hay_imputacion <- n_imputados > 0

    if (isTRUE(hay_discrepancia) || isTRUE(hay_imputacion)) {
      estado_archivo <- "semi-válido"
    }
  }

  # Crear estructura de resultados base
  result <- list(
    deviceID = device_id,
    date_processed = format(Sys.time(), "%d/%m/%Y %H:%M:%S"),
    position = position,
    corrupt = estado_archivo,   # ahora puede ser "corrupto" / "semi-válido" / "válido"
    n_faltantes = n_faltantes,  # paquetes que faltan (truncamiento)
    n_sobrantes = n_sobrantes,  # paquetes de más (cabecera mal grabada)
    n_imputados = n_imputados,  # paquetes presentes dañados (checksum fallido)
    n_declarados = n_declarados, # total declarados en cabecera
    n_observados = n_observados, # total observados
    starttime = NA,
    endtime = NA,
    duration_minutes = NA,
    duration_days = NA,
    acc = NA,
    dynrange_acc = NA,
    sf_acc = NA,
    gyro = NA,
    HR = NA,
    temp_body = NA,
    temp_ambient = NA
  )

  # Si readParmayMatrix devolvió objeto pero la estructura no es válida, avisar
  if (!is.null(bin_data) && !estructura_ok) {
    message("⚠️ El archivo se leyó pero la estructura de datos es incompleta (sin acelerometría utilizable): se marca como corrupto.")
  }

  # Si el archivo es legible (válido o semi-válido), extraer información
  if (!is.null(bin_data) && estructura_ok) {
    result$sf_acc <- bin_data$header$sf
    result$dynrange_acc <- bin_data$header$acc_dynrange

    # Calcular tiempos del registro.
    # Fuente principal: QClog$start / QClog$end (tiempos por bloque, eficiente y
    # disponible aunque el archivo no tenga acelerometría). Fallback: $data$time.
    qstart <- tryCatch(bin_data$QClog$start, error = function(e) NULL)
    qend   <- tryCatch(bin_data$QClog$end,   error = function(e) NULL)
    if (!is.null(qstart) && !is.null(qend) &&
        length(qstart) > 0 && length(qend) > 0 &&
        any(is.finite(qstart)) && any(is.finite(qend))) {
      result$starttime <- as.POSIXct(min(qstart, na.rm = TRUE), origin = "1970-01-01", tz = CONFIG$timezone)
      result$endtime   <- as.POSIXct(max(qend,   na.rm = TRUE), origin = "1970-01-01", tz = CONFIG$timezone)
    } else {
      # Fallback: usar la columna de tiempo de los datos
      times <- as.POSIXct(bin_data$data$time, origin = "1970-01-01", tz = CONFIG$timezone)
      result$starttime <- min(times, na.rm = TRUE)
      result$endtime   <- max(times, na.rm = TRUE)
    }
    result$duration_minutes <- round(as.numeric(difftime(result$endtime, result$starttime, units = "mins")), 2)
    result$duration_days <- round(result$duration_minutes / (60 * 24), 2)

    # Verificar sensores disponibles y convertir a formato Sí/No
    data_cols <- names(bin_data$data)
    result$acc <- ifelse("acc_x" %in% data_cols && any(!is.na(bin_data$data$acc_x)), "Sí", "No")
    result$gyro <- ifelse("gyro_x" %in% data_cols && any(!is.na(bin_data$data$gyro_x)), "Sí", "No")
    result$temp_ambient <- ifelse("ambient_temp" %in% data_cols && any(!is.na(bin_data$data$ambient_temp)), "Sí", "No")
    result$temp_body <- ifelse("bodySurface_temp" %in% data_cols && any(!is.na(bin_data$data$bodySurface_temp)), "Sí", "No")
    result$HR <- ifelse("hr_raw" %in% data_cols && any(!is.na(bin_data$data$hr_raw)), "Sí", "No")
  }

  return(list(result = result, data = bin_data))
}

#' Generar mensajes de resumen según las condiciones del archivo (esto es lo que se incluirá en los informes individuales)
generate_summary_messages <- function(result_data, base_dirs = NULL, participant_code = NULL) {
  messages <- c()

  if (result_data$corrupt == "corrupto") {
    # Mensaje para archivos corruptos
    messages <- c(messages, paste0(
      "<div style='color: red; font-size: 18px; font-weight: bold; margin: 20px 0;'>",
      "❌ Archivo Corrupto</div>",
      "<p>No se pudo extraer información del archivo porque está completamente dañado o corrupto.</p>",
      "<div style='background-color: #f0f0f0; padding: 15px; border-left: 4px solid #007bff; margin: 15px 0;'>",
      "<strong>📌 Nota importante:</strong><br>",
      "😥 Si ha restaurado la tarjeta del acelerómetro (es decir, ha presionado 'Restore TF Card' en la App Matrix al programar y colocar el mismo dispositivo en otro participante), no es posible intentar recuperar el archivo.<br><br>",
      "🤞 Si NO ha restaurado la tarjeta del acelerómetro, el archivo original de este participante aún estará en el acelerómetro y existe la posibilidad de que ahí no esté corrupto. En ese caso, se recomienda seguir los pasos descritos a continuación.",
      "</div>",
      "<div style='background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 15px 0;'>",
      "<strong>🛠️ Pasos recomendados:</strong><br>",
      "💻🔌 <strong>1.</strong> Conecte el acelerómetro a su PC mediante el cable USB.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;⚠️ <em>Recuerde: Nunca modifique el nombre del archivo en el dispositivo.</em><br>",
      "💾 <strong>2.</strong> Copie el archivo desde el dispositivo y péguelo en la carpeta 'archivos bin'. Espere a que se complete la transferencia del archivo. <br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;🔒 <em>Revise que el cable está conectado correctamente y de manera estable en ambos extremos (acelerómetro y puerto usb) y asegúrese de que el acelerómetro está situado en una superficie estable para evitar desconexiones durante la transferencia.</em><br>",
      "✏️ <strong>3.</strong> Renombre el archivo otorgándole el código que corresponda.<br>",
      "⚙️ <strong>4.</strong> Ejecute nuevamente el procesado para ese archivo.<br>",
      "🕵️️️️ <strong>5.</strong> Revise el informe individual generado (habrá sustituido al anterior) y compruebe si el archivo se calificó de nuevo como 'corrupto' o, por el contrario, resultó 'válido'.",
      "</div>"
    ))
    return(messages)
  }

  # Mensajes para archivos no corruptos

  # Nota informativa si el archivo es SEMI-VÁLIDO (se leyó bien, pero se eliminaron
  # algunos paquetes corruptos durante la decodificación). Aparece la primera.
  if (result_data$corrupt == "semi-válido") {
    n_falt <- ifelse(is.null(result_data$n_faltantes), 0, result_data$n_faltantes)
    n_sobr <- ifelse(is.null(result_data$n_sobrantes), 0, result_data$n_sobrantes)
    n_imp  <- ifelse(is.null(result_data$n_imputados), 0, result_data$n_imputados)

    # Describir cada incidencia detectada y su consejo correspondiente.
    detalle <- c()
    consejo <- c()
    if (n_falt > 0) {
      detalle <- c(detalle, paste0(
        "el archivo contiene menos paquetes de los que declara su cabecera (faltan ", n_falt,
        " paquetes), lo que suele indicar que la transferencia se interrumpió y parte del registro no llegó a copiarse"))
      consejo <- c(consejo, paste0(
        "transfiera el archivo de forma estable: conecte el acelerómetro directamente al ordenador (mejor que a ",
        "través de un hub o concentrador USB), colocado sobre una superficie estable, y espere a que la copia termine por ",
        "completo antes de desconectar el dispositivo"))
    }
    if (n_sobr > 0) {
      detalle <- c(detalle, paste0(
        "el archivo contiene más paquetes de los que declara su cabecera (", n_sobr,
        " de más), lo que indica que el dispositivo anotó de forma incorrecta el recuento de paquetes en la ",
        "cabecera del archivo; los datos del registro están completos y son analizables"))
      consejo <- c(consejo, paste0(
        "no es necesaria ninguna acción, ya que el archivo es válido y contiene todos los datos; si esta incidencia ",
        "se repitiera con frecuencia en un mismo dispositivo, conviene revisarlo"))
    }
    if (n_imp > 0) {
      detalle <- c(detalle, paste0(
        "se detectaron ", n_imp, " paquete(s) con datos dañados que se han rellenado automáticamente con valores ",
        "neutros; esto puede deberse a un fallo puntual durante la grabación (por ejemplo, un golpe o una vibración ",
        "fuerte que afecte momentáneamente al sensor) o a una incidencia durante la transferencia"))
      consejo <- c(consejo, paste0(
        "se trata de una incidencia menor y puntual, no de un fallo grave ni permanente del dispositivo; no conviene ",
        "indicar a los participantes que se retiren el acelerómetro en las situaciones que puedan provocar estos fallos ",
        "temporales (por ejemplo, durante la práctica de deporte), ya que perderíamos datos de esos periodos. Basta con ",
        "pedirles que procuren tratar el acelerómetro con el mayor cuidado posible durante su uso diario. Asimismo, los ",
        "investigadores deben mantener el máximo cuidado con el dispositivo y con el archivo de datos en todas las fases ",
        "(programación, colocación, retirada, descarga del archivo, transferencia, etc.)"))
    }
    if (length(detalle) == 0) {
      detalle <- "se detectaron incidencias menores de integridad en algunos paquetes"
      consejo <- "transfiera y maneje el dispositivo con cuidado para evitarlo en próximas ocasiones"
    }
    detalle_txt <- paste(detalle, collapse = "; ")
    consejo_txt <- paste(unique(consejo), collapse = "; ")

    messages <- c(messages, paste0(
      "🟡 <strong>Archivo (semi) válido.</strong> El archivo se ha leído y analizado correctamente, pero ",
      detalle_txt, ". Los datos presentes son válidos y se han procesado con normalidad, por lo que las ",
      "comprobaciones de este informe son fiables. RECOMENDACIÓN: ", consejo_txt, "."
    ))
  }

  # Mensaje si datos de aceleración no registrados
  if (result_data$acc == "No") {
    messages <- c(messages, "🚨 No se programó el acelerómetro para recoger datos de acelerometría (sensor principal; datos básicos). De este archivo no se podrá obtener información sobre conductas de movimiento. Por favor, revise el protocolo y, en próximas ocasiones, active la sección acc/gyro y seleccione '6 Axis' en la App Matrix cuando programe los acelerómetros que vaya a colocar en la muñeca o en el muslo.")
  }


  # Mensajes relacionados con la duración
  duration_hours <- result_data$duration_minutes / 60

  if (!is.na(duration_hours)) {
    if (duration_hours < 1) {  # Menos de 1 hora (60 minutos)
      messages <- c(messages, "⛔ Se ha registrado un tiempo mucho menor de lo esperado. Si ha descargado varios archivos de este participante en esta posición, revise la duración del resto de archivos. Es posible que el archivo de este informe se originó por iniciar y parar el acelerómetro una primera vez, antes de programarlo definitivamente y colocarlo.")
    } else if (duration_hours >= 1 && result_data$duration_days < 6) {  # Entre 1 hora y 6 días
      messages <- c(messages, "⛔ Se ha registrado un tiempo menor de lo esperado. En próximas ocasiones asegúrese de que el acelerómetro tiene el 100% de carga (o, al menos, más del 95%) antes de programarlo y colocarlo.")
    }

    if (result_data$duration_days > 9) {  # Más de 9 días
      messages <- c(messages, "⛔ Se ha registrado un tiempo mayor de lo esperado. Es posible que no parara el registro con la app al recoger el acelerómetro y/o que no lo programara con los parámetros correctos en cuanto a frecuencia de muestreo y sensores activados.")
    }
  }

  # Mensaje sobre frecuencia de muestreo
  if (!is.na(result_data$sf_acc) && result_data$sf_acc != 25) {
    messages <- c(messages, paste0("🚫 Se ha registrado la señal a ", result_data$sf_acc, " Hz. Por favor, revise el protocolo y, en próximas ocasiones, seleccione 25HZ en la sección 'frequency' de la App Matrix cuando programe los acelerómetros que vaya a colocar en la muñeca o en el muslo."))
  }

  # Mensaje sobre giróscopo
  if (result_data$gyro == "No") {
    messages <- c(messages, "🚫 No se programó el acelerómetro para recoger datos del giróscopo. Por favor, revise el protocolo y, en próximas ocasiones, seleccione '6 Axis' en la sección 'acc/gyro' de la App Matrix cuando programe los acelerómetros que vaya a colocar en la muñeca o en el muslo.")
  }

  # Mensajes específicos para muñeca
  if (result_data$position == "muñeca") {
    if (result_data$HR == "No") {
      messages <- c(messages, "🚫 No se programó el acelerómetro de muñeca para recoger datos de frecuencia cardiaca. Por favor, revise el protocolo y, en próximas ocasiones, active 'heart' y seleccione 'realTime' en la App Matrix al programar los acelerómetros que vaya a colocar en la muñeca.")
    }
    if (result_data$temp_body == "No") {
      messages <- c(messages, "🚫 No se programó el acelerómetro de muñeca para recoger datos de temperatura corporal. Por favor, revise el protocolo y, en próximas ocasiones, active 'skinTemp' y seleccione '1min/times' en la App Matrix al programar los acelerómetros que vaya a colocar en la muñeca.")
    }
    if (result_data$temp_ambient == "No") {
      messages <- c(messages, "🚫 No se programó el acelerómetro de muñeca para recoger datos de temperatura ambiente. Por favor, revise el protocolo y, en próximas ocasiones, active 'envTemp' y seleccione '1min/times' en la App Matrix al programar los acelerómetros que vaya a colocar en la muñeca.")
    }
  }

  # Mensajes específicos para muslo
  if (result_data$position == "muslo") {
    if (result_data$HR == "Sí") {
      messages <- c(messages, "⚠️ Se programó el acelerómetro de muslo para recoger datos de frecuencia cardiaca. Por favor, revise el protocolo y, en próximas ocasiones, no active 'heart' en la App Matrix al programar los acelerómetros que vaya a colocar en el muslo.")
    }
    if (result_data$temp_body == "Sí") {
      messages <- c(messages, "⚠️ Se programó el acelerómetro de muslo para recoger datos de temperatura corporal. Por favor, revise el protocolo y, en próximas ocasiones, no active 'skinTemp' en la App Matrix al programar los acelerómetros que vaya a colocar en el muslo.")
    }
    if (result_data$temp_ambient == "Sí") {
      messages <- c(messages, "⚠️ Se programó el acelerómetro de muslo para recoger datos de temperatura ambiente. Por favor, revise el protocolo y, en próximas ocasiones, no active 'envTemp' en la App Matrix al programar los acelerómetros que vaya a colocar en el muslo.")
    }
  }

  # Mensaje si no se puede obtener la posición (VERSIÓN EXPANDIDA CON RUTAS REALES)
  if (result_data$position == "desconocida") {
    # Construir rutas reales
    base_dir <- if (!is.null(base_dirs$processed_base)) dirname(base_dirs$processed_base) else getwd()
    if (basename(base_dir) == "scripts") {
      base_dir <- dirname(base_dir)
    }

    participant_path <- if (!is.null(participant_code)) {
      file.path(base_dir, "archivos bin procesados", participant_code)
    } else {
      file.path(base_dir, "archivos bin procesados", "[código_participante]")
    }

    bin_path <- file.path(base_dir, "archivos bin")

    # Nombres de archivos reales
    current_filename <- result_data$deviceID
    file_name_no_ext <- tools::file_path_sans_ext(current_filename)
    html_filename <- paste0("informe_", file_name_no_ext, ".html")
    excel_filename <- if (!is.null(participant_code)) {
      paste0("Resultados_", participant_code, ".xlsx")
    } else {
      "Resultados_[código_participante].xlsx"
    }

    # Ejemplos de nomenclatura correcta basados en el archivo actual
    muslo_example <- paste0("MD ACEMI-", sub(".*-", "", file_name_no_ext), ".bin")
    muneca_example <- paste0("MD ACEMS-", sub(".*-", "", file_name_no_ext), ".bin")

    messages <- c(messages, paste0(
      "<div style='color: #cc6600; font-size: 18px; font-weight: bold; margin: 20px 0;'>",
      "❓ Posición del acelerómetro desconocida</div>",
      "<p>No se pudo determinar la posición del acelerómetro (muñeca o muslo) a partir del nombre del archivo '<strong>", current_filename, "</strong>'. Esto se debe a que el nombre del archivo no sigue la nomenclatura estándar esperada.</p>",
      "<div style='background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 15px 0;'>",
      "<strong>📌 Nomenclatura esperada:</strong><br>",
      "• Para <strong>muslo</strong>: el nombre del archivo debe contener 'ACEMI'<br>",
      "• Para <strong>muñeca</strong>: el nombre del archivo debe contener 'ACEMS'",
      "</div>",
      "<div style='background-color: #e7f3ff; padding: 15px; border-left: 4px solid #007bff; margin: 15px 0;'>",
      "<strong>🛠️ Pasos recomendados:</strong><br>",
      "📁 <strong>1.</strong> Acceda a la carpeta de resultados de ese participante, en 'archivos bin procesados'.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📂 <em>Ruta: ", gsub("\\\\", "/", participant_path), "</em><br><br>",
      "🗑️ <strong>2.</strong> Elimine el informe HTML creado en el primer procesamiento. Acceda al Excel individual y elimine la fila de resultados de ese archivo.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📝 <em>Elimine: ", html_filename, "</em><br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📊 <em>Edite: ", excel_filename, " (eliminar fila correspondiente)</em><br><br>",
      "✂️ <strong>3.</strong> Corte el archivo con la nomenclatura incorrecta y péguelo en la carpeta 'archivos bin'.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📦 <em>Mover desde: ", gsub("\\\\", "/", participant_path), "</em><br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📦 <em>Mover hacia: ", gsub("\\\\", "/", bin_path), "</em><br><br>",
      "✏️ <strong>4.</strong> Renombre el archivo en la carpeta 'archivos bin' dotándole del nombre correcto.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;🏷️ <em>Incluya 'ACEMI' para muslo o 'ACEMS' para muñeca en el nombre</em><br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📝 <em>Ejemplo para muslo: '", current_filename, "' → '", muslo_example, "'</em><br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📝 <em>Ejemplo para muñeca: '", current_filename, "' → '", muneca_example, "'</em><br><br>",
      "▶️ <strong>5.</strong> Vuelva a ejecutar verAC.bat.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;⚙️ <em>El archivo será reprocesado automáticamente con el nuevo nombre</em><br><br>",
      "🔍 <strong>6.</strong> Revise el informe individual HTML y el Excel individual de resultados para corroborar que la posición del acelerómetro ahora está correctamente definida.<br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;✅ <em>Verifique que la columna 'Posición' ya no muestra 'desconocida'</em><br>",
      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📋 <em>Compruebe que los sensores esperados coinciden con la posición detectada</em>",
      "</div>"
    ))
    return(messages)
  }

  return(messages)
}

#' Generar reporte HTML y PDF en directorio del participante
generate_html_report <- function(result_data, participant_dir, base_dirs = NULL, participant_code = NULL) {
  # Cambiar nombre a informe_nombrearchivo.html
  file_name_no_ext <- tools::file_path_sans_ext(result_data$deviceID)
  html_filename <- file.path(participant_dir, paste0("informe_", file_name_no_ext, ".html"))

  # Crear backup del HTML si existe y está configurado
  create_backup_if_needed(html_filename, "html")

  # Esperar a que el archivo HTML esté disponible
  tryCatch({
    wait_for_html(html_filename)
  }, error = function(e) {
    # Si es error de tiempo agotado, re-lanzar SIN mensaje adicional
    if (grepl("TIEMPO_AGOTADO", e$message)) {
      stop(e$message)
    }
    message("❌ Error esperando disponibilidad del HTML: ", e$message)
    return(list(html = FALSE, pdf = FALSE, html_error = e$message))
  })

  # Obtener nombre del archivo sin extensión
  file_name_no_ext <- tools::file_path_sans_ext(result_data$deviceID)

  # Función para obtener sensores esperados según la posición
  get_expected_sensors <- function(position) {
    if (position == "muñeca") {
      expected_text <- "Esperado en archivo de muñeca"
      list(
        text = expected_text,
        accelerometer = "✅ Sí",
        gyro = "✅ Sí",
        HR = "✅ Sí",
        temp_body = "✅ Sí",
        temp_ambient = "✅ Sí"
      )
    } else if (position == "muslo") {
      expected_text <- "Esperado en archivo de muslo"
      list(
        text = expected_text,
        accelerometer = "✅ Sí",
        gyro = "✅ Sí",
        HR = "❌ No",
        temp_body = "❌ No",
        temp_ambient = "❌ No"
      )
    } else {
      expected_text <- "Esperado (posición desconocida)"
      list(
        text = expected_text,
        accelerometer = "✅ Sí",
        gyro = "✅ Sí",
        HR = "❓ Desconocido",
        temp_body = "❓ Desconocido",
        temp_ambient = "❓ Desconocido"
      )
    }
  }

  expected_sensors <- get_expected_sensors(result_data$position)

  # Generar mensajes de resumen
  summary_messages <- generate_summary_messages(result_data, base_dirs, participant_code)
  summary_html <- if (length(summary_messages) > 0) {
    paste0("<ul style='list-style-type: none; padding-left: 0;'>",
           paste0("<li style='margin-bottom: 10px;'>", summary_messages, "</li>", collapse = ""),
           "</ul>")
  } else {
    "<p>✅ No se detectaron problemas en el archivo.</p>"
  }

  # Plantilla HTML con CSS para salto de página
  html_content <- paste0(
    "<!DOCTYPE html>",
    "<html><head>",
    "<title>Informe: ", file_name_no_ext, "</title>",
    "<style>",
    "body{font-family:Arial,sans-serif;margin:40px;} ",
    "table{border-collapse:collapse;width:100%;margin-bottom:20px;} ",
    "th,td{border:1px solid #ddd;padding:8px;text-align:left;} ",
    "th{background-color:#f2f2f2;} ",
    ".centered{text-align:center;} ",
    ".page-break{page-break-before:always;} ",
    ".title{font-family:'Georgia',serif;color:#0066CC;font-size:28px;font-weight:bold;text-shadow:1px 1px 2px rgba(0,0,0,0.1);}",
    "</style>",
    "</head><body>",
    "<h1 class='centered title'>📊 Informe de calidad del archivo de acelerometría ", file_name_no_ext, "</h1>",
    "<h2> ℹ️ Información General</h2>",
    "<table>",
    "<tr><td>Nombre del archivo</td><td>", result_data$deviceID, "</td></tr>",
    "<tr><td>Posición</td><td>", result_data$position, "</td></tr>",
    "<tr><td>Fecha de procesamiento</td><td>", result_data$date_processed, "</td></tr>",
    "<tr><td>Estado del archivo</td><td>", switch(result_data$corrupt,
        "corrupto"     = "❌ Corrupto",
        "semi-válido"  = "🟡 (semi) válido",
        "válido"       = "✅ Válido",
        "✅ Válido"), "</td></tr>",
    "</table>",

    if (result_data$corrupt %in% c("válido", "semi-válido")) {
      paste0(
        "<h2>⏱️ Información Temporal</h2>",
        "<table>",
        "<tr><td>Inicio de grabación</td><td>", format(result_data$starttime, "%d/%m/%Y %H:%M:%S"), "</td></tr>",
        "<tr><td>Fin de grabación</td><td>", format(result_data$endtime, "%d/%m/%Y %H:%M:%S"), "</td></tr>",
        "<tr><td>Duración total</td><td>", round(result_data$duration_minutes, 1), " minutos (", round(result_data$duration_days, 2), " días)</td></tr>",
        "</table>",
        "<h2>🔍 Información de sensores</h2>",
        "<table>",
        "<colgroup><col style='width:34%'><col style='width:36%'><col style='width:30%'></colgroup>",
        "<tr><th>Sensor</th><th>", expected_sensors$text, "</th><th>Registrado</th></tr>",
        "<tr><td>Acelerómetro</td><td>", expected_sensors$accelerometer, "</td><td>", ifelse(result_data$acc == "Sí", "✅ Sí", "❌ No"), "</td></tr>",
        "<tr><td>Rango dinámico</td><td> 8 g</td><td>", ifelse(is.na(result_data$dynrange_acc), "❌ N/D", paste0(result_data$dynrange_acc, " g")), "</td></tr>",
        "<tr><td>Frecuencia de muestreo</td><td> 25 Hz</td><td>", ifelse(is.na(result_data$sf_acc), "❌ N/D", paste0(result_data$sf_acc, " Hz")), "</td></tr>",
        "<tr><td>Giróscopo</td><td>", expected_sensors$gyro, "</td><td>", ifelse(result_data$gyro == "Sí", "✅ Sí", "❌ No"), "</td></tr>",
        "<tr><td>Frecuencia cardiaca</td><td>", expected_sensors$HR, "</td><td>", ifelse(result_data$HR == "Sí", "✅ Sí", "❌ No"), "</td></tr>",
        "<tr><td>Temperatura corporal</td><td>", expected_sensors$temp_body, "</td><td>", ifelse(result_data$temp_body == "Sí", "✅ Sí", "❌ No"), "</td></tr>",
        "<tr><td>Temperatura ambiente</td><td>", expected_sensors$temp_ambient, "</td><td>", ifelse(result_data$temp_ambient == "Sí", "✅ Sí", "❌ No"), "</td></tr>",
        "</table>"
      )
    } else "",

    "<h2 class='page-break'>📋 Resumen y recomendaciones</h2>",
    summary_html,

    "</body></html>"
  )

  tryCatch({
    # Guardar HTML con supresión de avisos
    suppressWarnings({
      writeLines(html_content, html_filename)
    })

    # Generar PDF solo si generar_pdf=TRUE (ver CONFIG al inicio de la sintaxis) y si webshot2 (chrome) está disponible
    if (isTRUE(CONFIG$generar_pdf) && requireNamespace("webshot2", quietly = TRUE)) {
      tryCatch({
        pdf_filename <- file.path(participant_dir, paste0("informe_", file_name_no_ext, ".pdf"))

        # Crear backup del PDF si existe y está configurado
        create_backup_if_needed(pdf_filename, "html")  # Los PDF van con la misma configuración que HTML

        wait_for_pdf(pdf_filename)

        invisible(capture.output({
          suppressMessages(suppressWarnings({
            webshot2::webshot(
              url = html_filename,
              file = pdf_filename,
              vwidth = 800,
              vheight = 1200,
              delay = 1,
              cliprect = "viewport"
            )
          }))
        }, type = c("output", "message")))

        if (file.exists(pdf_filename)) {
          return(list(html = TRUE, pdf = TRUE))
        } else {
          return(list(html = TRUE, pdf = FALSE, pdf_error = "PDF no generado"))
        }

      }, error = function(e) {
        if (grepl("TIEMPO_AGOTADO", e$message)) {
          stop(e$message)
        }
        if (grepl("chrome|chromium", tolower(e$message))) {
          return(list(html = TRUE, pdf = FALSE, pdf_error = "Chrome/Chromium no encontrado"))
        } else {
          return(list(html = TRUE, pdf = FALSE, pdf_error = paste("Error PDF:", e$message)))
        }
      })

    } else if (!isTRUE(CONFIG$generar_pdf)) {
      # No hacer nada, simplemente devolver que solo se generó HTML
      return(list(html = TRUE, pdf = FALSE))
    } else {
      return(list(html = TRUE, pdf = FALSE, pdf_error = "webshot2 no disponible"))
    }


  }, error = function(e) {
    # Si es error de tiempo agotado, re-lanzar SIN mensaje adicional
    if (grepl("TIEMPO_AGOTADO", e$message)) {
      stop(e$message)
    }
    message("❌ Error generando HTML para ", result_data$deviceID, ": ", e$message)
    return(list(html = FALSE, pdf = FALSE, html_error = e$message))
  })
}

#' Guardar Excel individual por participante en su carpeta
save_participant_excel <- function(participant_data, participant_dir, participant_code) {
  excel_filename <- file.path(participant_dir, paste0("Resultados_", participant_code, ".xlsx"))

  # Esperar a que el archivo Excel esté disponible
  suppressWarnings({
    tryCatch({
      wait_for_excel(excel_filename)
    }, error = function(e) {
      if (grepl("TIEMPO_AGOTADO", e$message)) {
        stop(e$message)
      }
      message("❌ Error esperando disponibilidad del Excel individual: ", e$message)
      stop(e$message)
    })
  })

  # Crear backup del Excel individual si está configurado
  create_backup_if_needed(excel_filename, "excel_individual")

  # Limpiar duplicados si está configurado
  clean_excel_duplicates(excel_filename, participant_data$deviceID, CONFIG$limpiar_duplicados_excel_individual)

  # Convertir los datos a data.table con el mismo formato que el Excel global
  ordered_data <- data.table::data.table(
    ID_archivo = participant_data$deviceID,
    Fecha_procesamiento = participant_data$date_processed,
    Posición = participant_data$position,
    ESTADO_ARCHIVO = estado_label(participant_data$corrupt),
    Hora_inicio_grabación = participant_data$starttime,
    Hora_fin_grabación = participant_data$endtime,
    Duración_grabación_min = participant_data$duration_minutes,
    Duración_grabación_días = participant_data$duration_days,
    Acelerómetro = participant_data$acc,
    Rango_dinámico_acc = participant_data$dynrange_acc,
    Frecuencia_muestreo = participant_data$sf_acc,
    Giróscopo = participant_data$gyro,
    Frecuencia_cardiaca = participant_data$HR,
    Temperatura_corporal = participant_data$temp_body,
    Temperatura_ambiente = participant_data$temp_ambient
  )
  ordered_data <- add_info_paquetes(ordered_data, participant_data)
  date_style <- openxlsx::createStyle(numFmt = "DD/MM/YYYY HH:MM:SS")

  header_style <- openxlsx::createStyle(
    fgFill = "#D3D3D3",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left",
    valign = "top",
    wrapText = TRUE
  )

  border_style <- openxlsx::createStyle(
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  corrupt_style <- openxlsx::createStyle(
    fgFill = "#800000",
    fontColour = "#FFFFFF",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  red_style <- openxlsx::createStyle(
    fgFill = "#FF6666",
    fontColour = "#000000",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  orange_style <- openxlsx::createStyle(
    fgFill = "#FFE5B3",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  yellow_style <- openxlsx::createStyle(
    fgFill = "#FFFF99",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  # Aplicar el mismo formato que el Excel global (versión simplificada)
  suppressWarnings({
    if (file.exists(excel_filename)) {
      # Archivo existe - añadir datos
      wb <- openxlsx::loadWorkbook(excel_filename)
      existing_data <- tryCatch(
        openxlsx::read.xlsx(wb, sheet = 1),
        error = function(e) NULL
      )
      start_row <- if (!is.null(existing_data)) nrow(existing_data) + 2 else 2
      write_headers <- is.null(existing_data)

      openxlsx::writeData(wb, sheet = 1, x = ordered_data,
                          startRow = start_row, colNames = write_headers)
    } else {
      # Crear archivo nuevo
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "Sheet1")
      openxlsx::writeData(wb, sheet = 1, x = ordered_data, startRow = 1, colNames = TRUE)
    }

    # Leer todos los datos para aplicar formato
    all_data <- openxlsx::read.xlsx(wb, sheet = 1)
    total_rows <- nrow(all_data) + 1

    # 1. Formato de encabezados
    openxlsx::addStyle(wb, sheet = 1, style = header_style,
                       cols = 1:ncol(all_data), rows = 1, gridExpand = TRUE)
    openxlsx::setRowHeights(wb, sheet = 1, rows = 1, heights = 30)

    # 2. Bordes a todas las celdas
    openxlsx::addStyle(wb, sheet = 1, style = border_style,
                       cols = 1:ncol(all_data), rows = 2:total_rows, gridExpand = TRUE, stack = TRUE)

    # 3. Formato a fechas
    date_cols <- c(2, 5, 6)
    openxlsx::addStyle(wb, sheet = 1, style = date_style,
                       cols = date_cols, rows = 2:total_rows, gridExpand = TRUE, stack = TRUE)

    # 4. APLICAR ESTILOS DE ALERTA (igual que Excel global)
    corrupt_rows <- list()
    red_rows <- list()
    orange_rows <- list()
    yellow_rows <- list()

    for (i in 1:nrow(all_data)) {
      row_num <- i + 1

      # FONDO GRANATE
      if (!is.na(all_data$ESTADO_ARCHIVO[i]) && all_data$ESTADO_ARCHIVO[i] == "corrupto") {
        openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 4, rows = row_num, stack = TRUE)
        corrupt_rows[[length(corrupt_rows) + 1]] <- list(row = row_num, col = 4)
      }
      # FONDO AMARILLO: semi-válido
      if (!is.na(all_data$ESTADO_ARCHIVO[i]) && all_data$ESTADO_ARCHIVO[i] == "(semi) válido") {
        openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 4, rows = row_num, stack = TRUE)
        yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 4)
      }

      if (!is.na(all_data$Acelerómetro[i]) && all_data$Acelerómetro[i] == "No") {
        openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 9, rows = row_num, stack = TRUE)
        corrupt_rows[[length(corrupt_rows) + 1]] <- list(row = row_num, col = 9)
      }

      # FONDOS ROJOS
      if (!is.na(all_data$Duración_grabación_min[i]) && all_data$Duración_grabación_min[i] < 60) {
        openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 7, rows = row_num, stack = TRUE)
        red_rows[[length(red_rows) + 1]] <- list(row = row_num, col = 7)
      }

      if (!is.na(all_data$Posición[i]) && all_data$Posición[i] == "desconocida") {
        openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 3, rows = row_num, stack = TRUE)
        red_rows[[length(red_rows) + 1]] <- list(row = row_num, col = 3)
      }

      # FONDOS NARANJAS
      if (!is.na(all_data$Duración_grabación_días[i]) && all_data$Duración_grabación_días[i] < 6) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 8, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 8)
      }

      if (!is.na(all_data$Rango_dinámico_acc[i]) && all_data$Rango_dinámico_acc[i] != 8) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 10, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 10)
      }

      if (!is.na(all_data$Frecuencia_muestreo[i]) && all_data$Frecuencia_muestreo[i] != 25) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 11, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 11)
      }

      if (!is.na(all_data$Giróscopo[i]) && all_data$Giróscopo[i] == "No") {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 12, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 12)
      }

      # Condiciones específicas según posición
      position <- all_data$Posición[i]

      if (!is.na(position)) {
        if (position == "muñeca") {
          if (!is.na(all_data$Frecuencia_cardiaca[i]) && all_data$Frecuencia_cardiaca[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 13, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 13)
          }
          if (!is.na(all_data$Temperatura_corporal[i]) && all_data$Temperatura_corporal[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 14, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 14)
          }
          if (!is.na(all_data$Temperatura_ambiente[i]) && all_data$Temperatura_ambiente[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 15, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 15)
          }
        } else if (position == "muslo") {
          if (!is.na(all_data$Frecuencia_cardiaca[i]) && all_data$Frecuencia_cardiaca[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 13, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 13)
          }
          if (!is.na(all_data$Temperatura_corporal[i]) && all_data$Temperatura_corporal[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 14, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 14)
          }
          if (!is.na(all_data$Temperatura_ambiente[i]) && all_data$Temperatura_ambiente[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 15, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 15)
          }
        }
      }
    }

    # 5. APLICAR INDICADORES EN COLUMNA A según prioridad
    corrupt_row_nums <- unique(sapply(corrupt_rows, function(x) x$row))
    red_row_nums <- unique(sapply(red_rows, function(x) x$row))
    orange_row_nums <- unique(sapply(orange_rows, function(x) x$row))
    yellow_row_nums <- unique(sapply(yellow_rows, function(x) x$row))

    if (length(corrupt_row_nums) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 1, rows = corrupt_row_nums, stack = TRUE)
    }

    red_only <- setdiff(red_row_nums, corrupt_row_nums)
    if (length(red_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 1, rows = red_only, stack = TRUE)
    }

    orange_only <- setdiff(orange_row_nums, c(corrupt_row_nums, red_row_nums))
    if (length(orange_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 1, rows = orange_only, stack = TRUE)
    }

    yellow_only <- setdiff(yellow_row_nums, c(corrupt_row_nums, red_row_nums, orange_row_nums))
    if (length(yellow_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 1, rows = yellow_only, stack = TRUE)
    }

    # 6. Ajustar ancho de columnas
    openxlsx::setColWidths(wb, sheet = 1, cols = 1, widths = 28)
    openxlsx::setColWidths(wb, sheet = 1, cols = 2, widths = 21)
    openxlsx::setColWidths(wb, sheet = 1, cols = 3, widths = 12)
    openxlsx::setColWidths(wb, sheet = 1, cols = 4, widths = 12)
    openxlsx::setColWidths(wb, sheet = 1, cols = 5, widths = 20)
    openxlsx::setColWidths(wb, sheet = 1, cols = 6, widths = 20)
    openxlsx::setColWidths(wb, sheet = 1, cols = 7, widths = 18)
    openxlsx::setColWidths(wb, sheet = 1, cols = 8, widths = 18)
    openxlsx::setColWidths(wb, sheet = 1, cols = 9, widths = 14)
    openxlsx::setColWidths(wb, sheet = 1, cols = 10, widths = 15)
    openxlsx::setColWidths(wb, sheet = 1, cols = 11, widths = 11)
    openxlsx::setColWidths(wb, sheet = 1, cols = 12, widths = 11)
    openxlsx::setColWidths(wb, sheet = 1, cols = 13, widths = 10.5)
    openxlsx::setColWidths(wb, sheet = 1, cols = 14, widths = 12)
    openxlsx::setColWidths(wb, sheet = 1, cols = 15, widths = 12)

    # 7. Filtro automático
    openxlsx::addFilter(wb, sheet = 1, rows = 1, cols = 1:ncol(all_data))

    openxlsx::saveWorkbook(wb, excel_filename, overwrite = TRUE)
    repair_openxlsx_package(excel_filename)
  })

  message("📊 Excel individual guardado: Resultados_", participant_code, ".xlsx")
}


#' Guardar datos en Excel GLOBAL con formato chachi (en carpeta Excel de resultados del procesado)
save_to_global_excel <- function(result_data, excel_path) {
  # Esperar a que el archivo Excel esté disponible (con supresión de avisos)
  suppressWarnings({
    tryCatch({
      wait_for_excel(excel_path)
    }, error = function(e) {
      # Si es error de tiempo agotado, re-lanzar SIN mensaje adicional
      if (grepl("TIEMPO_AGOTADO", e$message)) {
        stop(e$message)
      }
      message("❌ Error esperando disponibilidad del Excel: ", e$message)
      stop(e$message)
    })
  })

  # Crear backup del Excel global si está configurado
  create_backup_if_needed(excel_path, "excel_global")

  # Limpiar duplicados si está configurado
  clean_excel_duplicates(excel_path, result_data$deviceID, CONFIG$limpiar_duplicados_excel_global)

  # Convertir los datos a data.table con el orden exacto especificado y nombres en español )
  ordered_data <- data.table::data.table(
    ID_archivo = result_data$deviceID,
    Fecha_procesamiento = result_data$date_processed,
    Posición = result_data$position,
    ESTADO_ARCHIVO = estado_label(result_data$corrupt),
    Hora_inicio_grabación = result_data$starttime,
    Hora_fin_grabación = result_data$endtime,
    Duración_grabación_min = result_data$duration_minutes,
    Duración_grabación_días = result_data$duration_days,
    Acelerómetro = result_data$acc,  # Nueva columna I
    Rango_dinámico_acc = result_data$dynrange_acc,  # Ahora columna J
    Frecuencia_muestreo = result_data$sf_acc,  # Ahora columna K
    Giróscopo = result_data$gyro,  # Ahora columna L
    Frecuencia_cardiaca = result_data$HR,  # Ahora columna M
    Temperatura_corporal = result_data$temp_body,  # Ahora columna N
    Temperatura_ambiente = result_data$temp_ambient  # Ahora columna O
  )
  ordered_data <- add_info_paquetes(ordered_data, result_data)

  # Definir estilos de formato
  date_style <- openxlsx::createStyle(numFmt = "DD/MM/YYYY HH:MM:SS")

  # Estilo para encabezados (fondo gris + alineación izquierda + ajuste texto + alineación superior)
  header_style <- openxlsx::createStyle(
    fgFill = "#D3D3D3",  # Gris claro
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left",  # Alineación horizontal izquierda
    valign = "top",   # Alineación vertical superior
    wrapText = TRUE   # Ajustar texto
  )

  # Estilo para bordes de datos + alineación izquierda
  border_style <- openxlsx::createStyle(
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"  # Alineación horizontal izquierda
  )

  # Estilos para diferentes niveles de alerta
  # Estilo para el problema más grave: archivos corruptos y acelerómetro = No
  corrupt_style <- openxlsx::createStyle(
    fgFill = "#800000",  # Granate
    fontColour = "#FFFFFF",  # Letra blanca
    textDecoration = "bold",  # Negrita
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  red_style <- openxlsx::createStyle(
    fgFill = "#FF6666",  # Rojo medio
    fontColour = "#000000",  # Letra negra
    textDecoration = "bold",  # Negrita
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  orange_style <- openxlsx::createStyle(
    fgFill = "#FFE5B3",  # Naranja claro
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  yellow_style <- openxlsx::createStyle(
    fgFill = "#FFFF99",  # Amarillo claro
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    halign = "left"
  )

  # Suprimir avisos durante todo el proceso de Excel
  suppressWarnings({
    if (file.exists(excel_path)) {
      # Archivo existe - añadir datos
      wb <- openxlsx::loadWorkbook(excel_path)

      if ("Sheet1" %in% names(wb)) {
        existing_data <- tryCatch(
          openxlsx::read.xlsx(wb, sheet = 1),
          error = function(e) NULL
        )
        start_row <- if (!is.null(existing_data)) nrow(existing_data) + 2 else 2
        write_headers <- is.null(existing_data)
      } else {
        openxlsx::addWorksheet(wb, "Sheet1")
        start_row <- 1
        write_headers <- TRUE
      }

      openxlsx::writeData(wb, sheet = 1, x = ordered_data,
                          startRow = start_row, colNames = write_headers)

      # Blindaje de la cabecera: reescribir SIEMPRE la fila 1 con todos los nombres de
      # columna actuales. Así, aunque el Excel existente tuviera una cabecera antigua o
      # incompleta (p. ej. sin las últimas columnas de sensores), la fila 1 queda completa
      # y correcta. Se escribe la cabecera como una fila de nombres, sin tocar los datos.
      encabezados <- as.list(names(ordered_data))
      for (k in seq_along(encabezados)) {
        openxlsx::writeData(wb, sheet = 1, x = encabezados[[k]],
                            startRow = 1, startCol = k, colNames = FALSE)
      }

      # Leer todos los datos para obtener el rango total
      all_data <- openxlsx::read.xlsx(wb, sheet = 1)
      total_rows <- nrow(all_data) + 1  # +1 para incluir encabezados

    } else {
      # Crear archivo nuevo
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "Sheet1")
      openxlsx::writeData(wb, sheet = 1, x = ordered_data, startRow = 1, colNames = TRUE)
      start_row <- 2
      total_rows <- nrow(ordered_data) + 1  # +1 para incluir encabezados
      all_data <- as.data.frame(ordered_data)  # para el rango de columnas de los formatos
    }

    # APLICAR FORMATOS

    # 1. Formato de encabezados (fila 1, columnas A:O - ahora son 15 columnas)
    openxlsx::addStyle(wb, sheet = 1, style = header_style,
                       cols = 1:ncol(all_data), rows = 1, gridExpand = TRUE)

    # Establecer altura de fila para encabezados
    openxlsx::setRowHeights(wb, sheet = 1, rows = 1, heights = 30)

    # 2. Aplicar bordes a todas las celdas con datos (A:O desde fila 1 hasta última fila)
    openxlsx::addStyle(wb, sheet = 1, style = border_style,
                       cols = 1:ncol(all_data), rows = 2:total_rows, gridExpand = TRUE, stack = TRUE)

    # 3. Aplicar formato a fechas (columnas 2, 5, 6: Fecha_procesamiento, Hora_inicio_grabación, Hora_fin_grabación)
    date_cols <- c(2, 5, 6)
    openxlsx::addStyle(wb, sheet = 1, style = date_style,
                       cols = date_cols, rows = 2:total_rows, gridExpand = TRUE, stack = TRUE)

    # 4. APLICAR ESTILOS DE ALERTA BASADOS EN CONDICIONES
    # Leer los datos para identificar filas que cumplen condiciones
    all_data <- openxlsx::read.xlsx(wb, sheet = 1)

    # Variables para rastrear qué filas tienen qué tipo de alertas
    corrupt_rows <- list()
    red_rows <- list()
    orange_rows <- list()
    yellow_rows <- list()

    for (i in 1:nrow(all_data)) {
      row_num <- i + 1  # +1 porque los encabezados están en fila 1

      # FONDO GRANATE (más grave): Archivos corruptos Y Acelerómetro = "No"
      # Columna D (ESTADO_ARCHIVO): "corrupto"
      if (!is.na(all_data$ESTADO_ARCHIVO[i]) && all_data$ESTADO_ARCHIVO[i] == "corrupto") {
        openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 4, rows = row_num, stack = TRUE)
        corrupt_rows[[length(corrupt_rows) + 1]] <- list(row = row_num, col = 4)
      }
      # FONDO AMARILLO: semi-válido (Columna D)
      if (!is.na(all_data$ESTADO_ARCHIVO[i]) && all_data$ESTADO_ARCHIVO[i] == "(semi) válido") {
        openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 4, rows = row_num, stack = TRUE)
        yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 4)
      }

      # Columna I (Acelerómetro): "No"
      if (!is.na(all_data$Acelerómetro[i]) && all_data$Acelerómetro[i] == "No") {
        openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 9, rows = row_num, stack = TRUE)
        corrupt_rows[[length(corrupt_rows) + 1]] <- list(row = row_num, col = 9)
      }

      # FONDOS ROJOS (graves)
      # Columna G (Duración_grabación_min): < 60
      if (!is.na(all_data$Duración_grabación_min[i]) && all_data$Duración_grabación_min[i] < 60) {
        openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 7, rows = row_num, stack = TRUE)
        red_rows[[length(red_rows) + 1]] <- list(row = row_num, col = 7)
      }

      # Columna C (Posición): "desconocida"
      if (!is.na(all_data$Posición[i]) && all_data$Posición[i] == "desconocida") {
        openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 3, rows = row_num, stack = TRUE)
        red_rows[[length(red_rows) + 1]] <- list(row = row_num, col = 3)
      }

      # FONDOS NARANJAS (advertencias)
      # Columna H (Duración_grabación_días): < 6
      if (!is.na(all_data$Duración_grabación_días[i]) && all_data$Duración_grabación_días[i] < 6) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 8, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 8)
      }
      # Columna H (Duración_grabación_días): > 9 (registro más largo de lo esperado) -> amarillo
      if (!is.na(all_data$Duración_grabación_días[i]) && all_data$Duración_grabación_días[i] > 9) {
        openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 8, rows = row_num, stack = TRUE)
        yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 8)
      }

      # Columna J (Rango_dinámico_acc): ≠ 8 (si no está vacía)
      if (!is.na(all_data$Rango_dinámico_acc[i]) && all_data$Rango_dinámico_acc[i] != 8) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 10, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 10)
      }

      # Columna K (Frecuencia_muestreo): ≠ 25 (si no está vacía)
      if (!is.na(all_data$Frecuencia_muestreo[i]) && all_data$Frecuencia_muestreo[i] != 25) {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 11, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 11)
      }

      # Columna L (Giróscopo): "No"
      if (!is.na(all_data$Giróscopo[i]) && all_data$Giróscopo[i] == "No") {
        openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 12, rows = row_num, stack = TRUE)
        orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 12)
      }

      # Condiciones específicas según posición
      position <- all_data$Posición[i]

      if (!is.na(position)) {
        if (position == "muñeca") {
          if (!is.na(all_data$Frecuencia_cardiaca[i]) && all_data$Frecuencia_cardiaca[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 13, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 13)
          }
          if (!is.na(all_data$Temperatura_corporal[i]) && all_data$Temperatura_corporal[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 14, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 14)
          }
          if (!is.na(all_data$Temperatura_ambiente[i]) && all_data$Temperatura_ambiente[i] == "No") {
            openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 15, rows = row_num, stack = TRUE)
            orange_rows[[length(orange_rows) + 1]] <- list(row = row_num, col = 15)
          }

        } else if (position == "muslo") {
          # Si Posición = "muslo": columnas M, N, O con "Sí" → AMARILLO
          if (!is.na(all_data$Frecuencia_cardiaca[i]) && all_data$Frecuencia_cardiaca[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 13, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 13)
          }
          if (!is.na(all_data$Temperatura_corporal[i]) && all_data$Temperatura_corporal[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 14, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 14)
          }
          if (!is.na(all_data$Temperatura_ambiente[i]) && all_data$Temperatura_ambiente[i] == "Sí") {
            openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 15, rows = row_num, stack = TRUE)
            yellow_rows[[length(yellow_rows) + 1]] <- list(row = row_num, col = 15)
          }
        }
      }
    }

    # 5. APLICAR INDICADORES EN COLUMNA A (deviceID) según prioridad
    # Crear vectores de filas únicas para cada tipo de alerta
    corrupt_row_nums <- unique(sapply(corrupt_rows, function(x) x$row))
    red_row_nums <- unique(sapply(red_rows, function(x) x$row))
    orange_row_nums <- unique(sapply(orange_rows, function(x) x$row))
    yellow_row_nums <- unique(sapply(yellow_rows, function(x) x$row))

    # Aplicar indicadores por orden de prioridad (granate prevalece sobre todo)
    if (length(corrupt_row_nums) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = corrupt_style, cols = 1, rows = corrupt_row_nums, stack = TRUE)
    }

    # Rojo solo en filas que no tienen granate
    red_only <- setdiff(red_row_nums, corrupt_row_nums)
    if (length(red_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = red_style, cols = 1, rows = red_only, stack = TRUE)
    }

    # Naranja solo en filas que no tienen granate ni rojo
    orange_only <- setdiff(orange_row_nums, c(corrupt_row_nums, red_row_nums))
    if (length(orange_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = orange_style, cols = 1, rows = orange_only, stack = TRUE)
    }

    # Amarillo solo en filas que no tienen granate, rojo ni naranja
    yellow_only <- setdiff(yellow_row_nums, c(corrupt_row_nums, red_row_nums, orange_row_nums))
    if (length(yellow_only) > 0) {
      openxlsx::addStyle(wb, sheet = 1, style = yellow_style, cols = 1, rows = yellow_only, stack = TRUE)
    }

    # 6. Ajustar ancho de columnas con anchos específicos
    openxlsx::setColWidths(wb, sheet = 1, cols = 1, widths = 28)    # ID_archivo
    openxlsx::setColWidths(wb, sheet = 1, cols = 2, widths = 21)    # Fecha_procesamiento
    openxlsx::setColWidths(wb, sheet = 1, cols = 3, widths = 12)    # Posición
    openxlsx::setColWidths(wb, sheet = 1, cols = 4, widths = 12)    # CORRUPTO
    openxlsx::setColWidths(wb, sheet = 1, cols = 5, widths = 20)    # Hora_inicio_grabación
    openxlsx::setColWidths(wb, sheet = 1, cols = 6, widths = 20)    # Hora_fin_grabación
    openxlsx::setColWidths(wb, sheet = 1, cols = 7, widths = 18)    # Duración_grabación_min
    openxlsx::setColWidths(wb, sheet = 1, cols = 8, widths = 18)    # Duración_grabación_días
    openxlsx::setColWidths(wb, sheet = 1, cols = 9, widths = 14)    # Acelerómetro (NUEVA)
    openxlsx::setColWidths(wb, sheet = 1, cols = 10, widths = 15)   # Rango_dinámico_acc
    openxlsx::setColWidths(wb, sheet = 1, cols = 11, widths = 11)   # Frecuencia_muestreo
    openxlsx::setColWidths(wb, sheet = 1, cols = 12, widths = 11)   # Giróscopo
    openxlsx::setColWidths(wb, sheet = 1, cols = 13, widths = 10.5) # Frecuencia_cardiaca
    openxlsx::setColWidths(wb, sheet = 1, cols = 14, widths = 12)   # Temperatura_corporal
    openxlsx::setColWidths(wb, sheet = 1, cols = 15, widths = 12)   # Temperatura_ambiente

    # 7. Añadir filtro automático a la fila 1 (ahora 18 columnas)
    openxlsx::addFilter(wb, sheet = 1, rows = 1, cols = 1:ncol(all_data))

    # Guardar el archivo
    openxlsx::saveWorkbook(wb, excel_path, overwrite = TRUE)
    repair_openxlsx_package(excel_path)
  })
}


#' Guardar archivo CSV si está configurado
should_generate_csv_for_position <- function(position) {
  return(position %in% CONFIG$csv_positions)
}

save_csv_if_needed <- function(bin_data, original_path, csv_dir) {
  if (!CONFIG$guardar_csv || is.null(bin_data)) return()

  # Determinar la posición del dispositivo
  position <- get_device_position(basename(original_path))

  # Verificar si se debe generar CSV para esta posición
  if (!should_generate_csv_for_position(position)) {
    return()
  }

  csv_filename <- file.path(csv_dir, sub("\\.bin$", ".csv", basename(original_path), ignore.case = TRUE))

  if (file.exists(csv_filename) && !CONFIG$sobreescribir_csv) {
    message("⚠️ El CSV bruto ya existe (no se sobrescribe): ", basename(csv_filename))
    return()
  }

  tryCatch({
    # Extraer datos de la salida que se van a añadir a la cabecera del csv
    deviceID <- tools::file_path_sans_ext(basename(original_path))  # Se tomará como ID el nombre del archivo
    SF <- bin_data$header$sf  # Frecuencia de muestreo
    starttime <- bin_data$header$starttime  # Tiempo de inicio (POSIXct)
    startTimeISO <- format(starttime, "%Y%m%dT%H%M%OS3") # Tiempo de inicio transformado de POSIXct a formato ISO8601 (es lo apropiado para el ActiPASS)

    # Modificar los axis (aceleración) para convención ActiPASS/Axivity (muslo).
    # Se aplica SIEMPRE a los CSV de muslo. ActiPASS detecta y corrige automáticamente
    # la orientación (upright/inverted) durante su análisis.
    rot <- rotate_axes_basic(bin_data)
    if (is.null(rot)) {
      message("❌ No se pudo aplicar la transformación de ejes (datos de acelerometría no disponibles).")
      return()
    }
    acc_x <- rot$x
    acc_y <- rot$y
    acc_z <- rot$z


    # Crear la cabecera del archivo CSV (necesaria para que el ActiPASS lea/interprete correctamente el csv)
    header_lines <- c(
      paste0("ID=", deviceID),
      paste0("SF=", SF),
      paste0("START=", startTimeISO),
      "x,y,z"
    )

    suppressWarnings({
      writeLines(header_lines, csv_filename)
    })

    # Crear los datos de aceleración (x, y, z)
    data_to_save <- data.frame(acc_x, acc_y, acc_z)

    # Guardar CSV (cabecera + datos de aceleración x, y, z)
    suppressWarnings({
      write.table(data_to_save, file = csv_filename, sep = ",", col.names = FALSE, row.names = FALSE, append = TRUE)
    })

    message("✅ CSV bruto guardado: ", basename(csv_filename))

  }, error = function(e) {
    message("❌ Error guardando CSV bruto: ", e$message)
  })
}

#' Mover archivo procesado a carpeta del participante
move_processed_file <- function(source_path, participant_dir) {
  dest_path <- file.path(participant_dir, basename(source_path))

  # Crear backup del archivo .bin si ya existe y está configurado
  create_backup_if_needed(dest_path, "bin")

  if (file.exists(dest_path)) {
    file.remove(dest_path)
  }

  if (file.rename(source_path, dest_path)) {
    return(TRUE)
  } else {
    message("❌ Error moviendo archivo: ", basename(source_path))
    return(FALSE)
  }
}


# -----------------------------------------------------------------------------
# 2b. GESTIÓN DE PAQUETES: INSTALACIÓN LOCAL DESDE BINARIOS .zip (sin Internet)
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Funcion: localizar la carpeta scripts/paquetesR a partir de la ubicacion real
# del script en ejecucion (robusta a ejecutar desde la carpeta padre o desde scripts)
# -----------------------------------------------------------------------------
get_paquetes_dir <- function() {
  args <- commandArgs(trailingOnly = FALSE)
  script_path <- sub("^--file=", "", grep("^--file=", args, value = TRUE))
  if (length(script_path) == 0 || script_path == "") {
    script_dir <- getwd()
  } else {
    script_dir <- dirname(normalizePath(script_path, winslash = "/", mustWork = FALSE))
  }
  # El script principal vive en scripts/ ; los .zip estan en scripts/paquetesR/
  candidate <- file.path(script_dir, "paquetesR")
  if (dir.exists(candidate)) return(candidate)
  # Por si se ejecuta desde la carpeta padre:
  candidate2 <- file.path(script_dir, "scripts", "paquetesR")
  if (dir.exists(candidate2)) return(candidate2)
  return(candidate)  # se devolvera aunque no exista; se valida despues
}

# -----------------------------------------------------------------------------
# Funcion principal de instalacion/carga local de paquetes
# -----------------------------------------------------------------------------
setup_packages_local <- function() {
  # NOTA: La INSTALACIÓN de paquetes la realiza 'Instalador_verAC.bat' (una vez, por TI).
  # Aquí, el procesador SOLO VERIFICA que el entorno está disponible y carga los paquetes.
  # Si falta algo, NO intenta instalar: indica ejecutar el instalador y sale con código 2.
  message("🔍 Verificando entorno (paquetes ya instalados por el instalador)...")

  core_min <- list("data.table" = "1.17.0", "openxlsx" = "4.2.8",
                   "GGIRread" = "1.0.8", "zip" = "3.0.0")

  is_ok <- function(pkg) {
    minv <- core_min[[pkg]]
    if (!requireNamespace(pkg, quietly = TRUE)) return(FALSE)
    if (is.null(minv)) return(TRUE)
    utils::compareVersion(as.character(packageVersion(pkg)), minv) >= 0
  }

  faltan <- names(core_min)[!vapply(names(core_min), is_ok, logical(1))]
  if (length(faltan) > 0) {
    message("❌ Faltan paquetes imprescindibles o con versión insuficiente: ",
            paste(faltan, collapse = ", "))
    message("   El entorno no está completo. Pida a su servicio de informática que ejecute")
    message("   'Instalador_verAC.bat' una vez para instalar R y los paquetes. Después, reintente.")
    quit(save = "no", status = 2)
  }

  message("🔄 Cargando paquetes...")
  for (pkg in names(core_min)) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      message("❌ No se pudo cargar el paquete ", pkg, ". Ejecute 'Instalador_verAC.bat'.")
      quit(save = "no", status = 2)
    }
  }
  message("✅ Todos los paquetes imprescindibles están disponibles y cargados")
  message("================================================================================")

  # ------------------------------------------------------------------------
  # CAPA 1 anti-falso-corrupto: verificar que GGIRread está OPERATIVO antes
  # de procesar. Si la función de lectura o sus dependencias no funcionan,
  # detener aquí con un mensaje claro, en vez de procesar y marcar TODOS los
  # archivos como "corruptos" (falso positivo masivo).
  # ------------------------------------------------------------------------
  if (!exists("readParmayMatrix", where = asNamespace("GGIRread"), inherits = FALSE) &&
      !is.function(tryCatch(GGIRread::readParmayMatrix, error = function(e) NULL))) {
    message("❌ ERROR CRÍTICO: la función 'readParmayMatrix' de GGIRread no está disponible.")
    message("   Esto indica un problema con GGIRread o sus dependencias, NO con los archivos .bin.")
    message("   Ejecute 'Instalador_verAC.bat' para reinstalar el entorno.")
    quit(save = "no", status = 2)
  }
}


# -----------------------------------------------------------------------------
# 3. PROGRAMA PRINCIPAL
# -----------------------------------------------------------------------------

main <- function() {
  # 1. FASE DE VERIFICACIÓN E INSTALACIÓN DE PAQUETES (instalación LOCAL desde .zip)
  setup_packages_local()


  # 3. CONTINUAR CON EL PROCESAMIENTO
  message("🔍 Verificando carpeta de archivos...")
  message("")

  # Configurar directorios base
  dirs <- setup_directories()

  # Buscar archivos BIN
  bin_files <- list.files(dirs$bin, pattern = "\\.bin$",
                          full.names = TRUE, ignore.case = TRUE)

  if (length(bin_files) == 0) {
    cat("Error:\n")
    cat("  ❗ No se encontraron archivos .bin en: ", gsub("\\\\", "/", dirs$bin),". Ejecución interrumpida.\n")
    quit(status = 1)
  }

  message("📁 Encontrados ", length(bin_files), " archivos .bin para procesar")
  message("--------------------------------------------------------------------------------")



  # Archivo Excel global de resultados
  excel_file <- file.path(dirs$excel_results, "RESULTADOS_COMPROBACION_DE_ARCHIVOS_ALL.xlsx")

  # Variables para contar PDFs generados
  pdf_count <- 0

  # Variables para control de errores críticos
  errores_criticos <- FALSE
  archivos_procesados <- 0
  archivos_error_entorno <- 0
  tipo_error <- ""

  # Contadores por categoría para el resumen final
  n_validos <- 0       # válidos o semi-válidos (informe generado, archivo utilizable)
  n_corruptos <- 0     # corruptos (datos dañados o sin estructura)
  n_pos_desconocida <- 0  # informativo: archivos cuya posición no se pudo determinar
  n_id_desconocido <- 0   # informativo: archivos cuyo ID/código no se pudo identificar

  # Variable para rastrear datos de participantes (para Excel individual)
  participant_data_tracker <- list()

  # Procesar cada archivo
  for (i in seq_along(bin_files)) {
    bin_file <- bin_files[i]

    # Si ya hay errores críticos, salir del bucle
    if (errores_criticos) {
      break
    }

    # Extraer código de participante
    participant_code <- get_participant_code(basename(bin_file))
    participant_dir <- file.path(dirs$processed_base, participant_code)

    # Verificar si el archivo ya fue procesado
    already_processed <- is_file_already_processed(basename(bin_file), participant_dir)

    if (already_processed && !CONFIG$reprocesar_existentes) {
      message("⏭️ Saltando archivo ya procesado: ", basename(bin_file))
      next
    }

    if (already_processed) {
      message("🔄 Reprocesando archivo existente: ", basename(bin_file))

    } else {
      message("🔄 Procesando: ", basename(bin_file), " (", get_device_position(basename(bin_file)), ")")
    }

    tryCatch({

      message("👤 Código de participante: ", participant_code)
      if (identical(participant_code, "desconocido")) {
        n_id_desconocido <- n_id_desconocido + 1
      }

      # Configurar directorio del participante
      participant_dir <- setup_participant_directory(dirs, participant_code)

      # Procesar archivo BIN
      processed <- process_bin_file(bin_file)

      # CASO ESPECIAL: problema de herramienta/entorno (no es archivo corrupto).
      # No se genera informe ni Excel; se avisa y se cuenta aparte. El archivo NO
      # se mueve (se deja en 'archivos bin' para reintentar tras revisar la instalación).
      if (is.null(processed$result)) {
        message("⏭️ Archivo OMITIDO por error de entorno (no corrupto): ", basename(bin_file))
        message("   Revise la instalación (Instalador_verAC.bat) y reintente este archivo.")
        message("--------------------------------------------------------------------------------")
        archivos_error_entorno <- archivos_error_entorno + 1
        next
      }

      if (processed$result$corrupt == "corrupto") {
        message("❌ Archivo corrupto: ", basename(bin_file))
        n_corruptos <- n_corruptos + 1
      } else if (processed$result$corrupt == "semi-válido") {
        message("🟡 Archivo (semi) válido (truncamiento y/o paquetes imputados): ", basename(bin_file))
        n_validos <- n_validos + 1
      } else {
        message("✅ Datos extraídos correctamente del archivo: ", basename(bin_file))
        n_validos <- n_validos + 1
      }
      # Contador informativo de posición desconocida
      if (!is.null(processed$result$position) && processed$result$position == "desconocida") {
        n_pos_desconocida <- n_pos_desconocida + 1
      }

      # Generar reporte HTML y PDF (si está habilitado) en carpeta del participante
      report_result <- generate_html_report(processed$result, participant_dir, dirs, participant_code)

      if (report_result$html) {
        message("📝 Reporte HTML generado en carpeta del participante: informe_", tools::file_path_sans_ext(basename(bin_file)), ".html")

        if (report_result$pdf) {
          message("📄 Reporte PDF generado en carpeta del participante: informe_", tools::file_path_sans_ext(basename(bin_file)), ".pdf")
          pdf_count <- pdf_count + 1
        } else if (!is.null(report_result$pdf_error)) {
          if (report_result$pdf_error == "Chrome/Chromium no encontrado") {
            message("⚠️ PDF no generado: Chrome/Chromium no está instalado en el sistema")
          } else if (report_result$pdf_error == "webshot2 no disponible") {
            message("⚠️ PDF no generado: webshot2 no está disponible")
          } else {
            message("⚠️ PDF no generado: ", report_result$pdf_error)
          }
        }
      } else {
        message("❌ Error generando reporte HTML")
      }

      # Guardar en Excel individual del participante
      save_participant_excel(processed$result, participant_dir, participant_code)

      # Guardar en Excel GLOBAL
      save_to_global_excel(processed$result, excel_file)
      message("📊 Datos añadidos al Excel global 'RESULTADOS_COMPROBACION_DE_ARCHIVOS_ALL.xlsx'")


      # Guardar CSV si especificado
      save_csv_if_needed(processed$data, bin_file, dirs$csv)

      # Mover archivo procesado a carpeta del participante
      move_processed_file(bin_file, participant_dir)
      if (already_processed) {
        message("📦 Archivo ", basename(bin_file), " sobrescrito en carpeta del participante: ", participant_code)
      } else {
        message("📦 Archivo ", basename(bin_file), " movido a carpeta del participante: ", participant_code)
      }

      if (already_processed) {
        message("✔️ Reprocesamiento completado: ", basename(bin_file))
        message("--------------------------------------------------------------------------------")

      } else {
        message("✔️ Procesamiento completado: ", basename(bin_file))
        message("--------------------------------------------------------------------------------")

      }
      archivos_procesados <- archivos_procesados + 1

    }, error = function(e) {
      message("⚠️ Fallo al procesar ", basename(bin_file), " (NO es corrupto): ", e$message)
      message("   Este fallo es de la herramienta/guardado, NO del contenido del archivo.")
      message("   El archivo NO se descarta ni se marca como corrupto. Revise el entorno y reejecute.")
      message("   Si el problema persiste, contacte con: veronica.cabanas@uam.es")
      archivos_error_entorno <<- archivos_error_entorno + 1

      # Verificar si es un error crítico de tiempo agotado
      if (grepl("TIEMPO_AGOTADO", e$message)) {
        errores_criticos <<- TRUE

        # Determinar tipo de archivo
        if (grepl("TIEMPO_AGOTADO_EXCEL", e$message)) {
          tipo_error <<- "Excel"
        } else if (grepl("TIEMPO_AGOTADO_HTML", e$message)) {
          tipo_error <<- "HTML"
        } else if (grepl("TIEMPO_AGOTADO_PDF", e$message)) {
          tipo_error <<- "PDF"
        } else {
          tipo_error <<- "archivo"
        }

        message("⛔ Error crítico detectado. Deteniendo procesamiento.")
      }
    })
  }

  message("--------------------------------------------------------------------------------")

  if (errores_criticos) {
    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    message("⛔ PROCESAMIENTO INTERRUMPIDO")
    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    message("")
    message("📊 Archivos procesados antes de la interrupción: ", archivos_procesados, " de ", length(bin_files))
    message("⚠️ El procesamiento se detuvo porque hay un archivo ", tipo_error, " abierto que se tiene que sobrescribir.")
    message("")
    quit(status = 3)

  } else {
    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    message("📊 RESUMEN DEL PROCESAMIENTO")
    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    message("   Archivos leídos:                       ", length(bin_files))
    message("   Informes generados (válidos o semi):   ", n_validos)
    message("   Corruptos:                             ", n_corruptos)
    if (archivos_error_entorno > 0) {
      message("   Omitidos por error de entorno:         ", archivos_error_entorno)
    }
    message("   ----------------------------------------------------------------")
    message("   (informativo) Con posición desconocida: ", n_pos_desconocida)
    message("   (informativo) Con ID/código desconocido: ", n_id_desconocido)
    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    if (archivos_error_entorno > 0) {
      message("⛔ Nota sobre los OMITIDOS por error de entorno (NO corruptos):")
      message("   No se procesaron por un fallo de la herramienta/entorno (p.ej. un paquete o")
      message("   dependencia de R bloqueado), no porque estén dañados. Se han dejado en la")
      message("   carpeta 'archivos bin'. Revise la instalación (Instalador_verAC.bat) y vuelva")
      message("   a ejecutar verAC. Si el problema persiste, contacte con: veronica.cabanas@uam.es")
    }
    message("")

    # Mostrar resultados con rutas usando "/"
    excel_path_display <- gsub("\\\\", "/", excel_file)
    processed_path_display <- gsub("\\\\", "/", dirs$processed_base)

    message("📁 Resultados guardados en:")
    message("   📊 Excel global: ", excel_path_display)
    message("   📁 Archivos organizados por participante: ", processed_path_display)
    message("   📝 Informes HTML individuales: en cada carpeta de participante")
    message("   📊 Excel individual por participante: en cada carpeta de participante")

    # Mostrar información sobre backups si están activados
    backup_messages <- c()
    if (CONFIG$backup_archivos_bin) backup_messages <- c(backup_messages, "archivos .bin")
    if (CONFIG$backup_archivos_html) backup_messages <- c(backup_messages, "archivos .html")
    if (CONFIG$backup_excel_individual) backup_messages <- c(backup_messages, "Excel individuales")
    if (CONFIG$backup_excel_global) backup_messages <- c(backup_messages, "Excel global")

    if (length(backup_messages) > 0) {
      message("   🗄️ Backups activados para: ", paste(backup_messages, collapse = ", "))
    } else {
      message("   🗄️ Backups: desactivados para todos los tipos de archivo")
    }

    # Mostrar información sobre duplicados
    if (CONFIG$limpiar_duplicados_excel_global && CONFIG$limpiar_duplicados_excel_individual) {
      message("   🧹 Duplicados Excel: eliminados automáticamente al reprocesar (global e individuales)")
    } else if (CONFIG$limpiar_duplicados_excel_global) {
      message("   🧹 Duplicados Excel global: eliminados automáticamente al reprocesar")
      message("   📋 Historial Excel individuales: mantenido (sin eliminar duplicados)")
    } else if (CONFIG$limpiar_duplicados_excel_individual) {
      message("   📋 Historial Excel global: mantenido (sin eliminar duplicados)")
      message("   🧹 Duplicados Excel individuales: eliminados automáticamente al reprocesar")
    } else {
      message("   📋 Historial Excel: mantenido en ambos tipos (sin eliminar duplicados)")
    }

    # Mostrar información sobre PDFs (si generar_pdfs=TRUE)
    if (isTRUE(CONFIG$generar_pdf)) {
      if (pdf_count > 0) {
        message("   📄 Informes PDF: en cada carpeta de participante")
      } else {
        message("   📄 Informes PDF: No se generaron probablemente porque Chrome no está instalado en el sistema")
      }
    }

    message("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
  }
}

# Ejecutar el programa principal
if (!interactive()) {
  main()
}
