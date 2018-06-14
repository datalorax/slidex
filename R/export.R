#' Extract xml from pptx
#'
#' @param path Path to the Microsoft PowerPoint file
#' @param force If an 'assets' folder already exists in the current directory,
#'   (e.g., from a previous conversion) should it be overwritten? Defaults to
#'   \code{FALSE}.
#' @inheritParams create_yaml
#' @param writenotes Logical. If notes are in the original PPTX, should they be
#'   written out as their own .txt file in addition to being embedded in the
#'   slides? Defaults to \code{TRUE}.
#' @param out_dir Optional output directory for the folder containing the RMD
#'   and corresponding assets. Defaults to the current working directory.
#' @export
#'
convert_pptx <- function(path, author, title = NULL, sub = NULL,
                         date = Sys.Date(), theme = "default",
                         highlightStyle = "github", force = FALSE,
                         writenotes = TRUE, out_dir = NULL) {

  if(!file.exists(path)) {
    stop(paste0("Cannot find file ", basename(path), " in directory",
                "'", dirname(path), "'",
                ". ", "Note - file paths must be specified with the '.pptx'",
                "extension."))
  }
  xml <- extract_xml(path, force = force)
  on.exit(unlink(xml, recursive = TRUE), add = TRUE)
  folder <- gsub("\\.pptx", "", basename(path))

  if(!is.null(out_dir)) {
    if(file.exists(file.path(out_dir, folder)) & force == FALSE) {
      stop(
        paste0(
          "This function will create a new folder in this ",
          "directory with the folder of the PPTX, but a folder with this ",
          "folder already exists here. Please move/delete the folder, ",
          "specify a new output directory with `out = 'path'`, or rerun ",
          "with `force = TRUE` to force the function to overwrite the ",
          "existing folder and all the files within it."
          )
        )
      }
  }

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

  rmd <- file.path(folder, paste0(folder, ".Rmd"))

  sink_error <- tryCatch(
    sink_rmd(xml, rmd, slds, rels,
             title_sld, author, title, sub, date, theme,
             highlightStyle),
    error = function(e) e
  )

  if(!is.null(sink_error$message)) {
    unlink(folder, recursive = TRUE)
    stop(sink_error$message)
  }

  if(length(list.files(file.path(folder, "assets"))) == 0) {
    unlink(file.path(folder, "assets"), recursive = TRUE)
  }

  if(writenotes) {
    write_notes(xml)
  }
  if(!is.null(out_dir)) {
    file.copy(folder, out_dir, recursive = TRUE)
    unlink(folder, recursive = TRUE, force = TRUE)
    system(paste(Sys.getenv("R_BROWSER"), file.path(out_dir, rmd)))
  }
  else{
    system(paste(Sys.getenv("R_BROWSER"), rmd))
  }
}


