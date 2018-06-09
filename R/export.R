#' Extract xml from pptx
#'
#' @param path Path to the Microsoft PowerPoint file
#' @param force If an 'assets' folder already exists in the current directory,
#'   (e.g., from a previous conversion) should it be overwritten? Defaults to
#'   \code{FALSE}.
#' @inheritParams create_yaml
#' @export
#'
convert_pptx <- function(path, author, title = NULL, sub = NULL,
                         date = Sys.Date(), theme = "default",
                         highlightStyle = "github", force = FALSE) {

  xml <- paste0(gsub("\\.pptx", "", basename(path)), "_xml")
  on.exit(unlink(xml, recursive = TRUE), add = TRUE)

  if(!file.exists(path)) {
    stop(paste0("Cannot find file ", basename(path), " in directory",
                "'", dirname(path), "'",
                ". ", "Note - file paths must be specified with the '.pptx'",
                "extension."))
  }
  extract_xml(path, force = force)

  lang_return <- tryCatch(check_lang(xml), error = function(e) e)
  if(!is.null(lang_return$message)) {
    unlink("assets", recursive = TRUE)
    stop(lang_return$message)
  }

  slds <- import_slide_xml(xml)
  rels <- import_rel_xml(xml)

  title_sld <- slds[[1]]

  slds <- slds[-1]
  rels <- rels[-1]

  rmd <- paste0(gsub("_xml", "", xml), ".Rmd")

  sink_error <- tryCatch(
    sink_rmd(rmd, slds, rels,
             title_sld, author, title, sub, date, theme,
             highlightStyle),
    error = function(e) e
  )

  if(!is.null(sink_error$message)) {
    unlink("assets", recursive = TRUE)
    unlink(rmd, recursive = TRUE)
    stop(sink_error$message)
  }

  if(length(list.files("assets")) == 0) {
    unlink("assets", recursive = TRUE)
  }

  system(paste(Sys.getenv("R_BROWSER"), rmd))
}


