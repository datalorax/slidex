#' Check the Language of PPTX
#' 
#' @param xml_folder The folder containing all of the xml code from the pptx,
#' created from \code{\link{extract_xml}}.
#' @keywords internal
#' 
check_lang <- function(xml_folder) {
  pres_xml <- read_xml(file.path(xml_folder, "ppt", "presentation.xml"))

  lang <- xml_find_all(pres_xml, "//p:defaultTextStyle/a:defPPr/a:defRPr") %>%
    map_chr(~xml_attr(., "lang"))

  if(any(lang != "en-US")) {
    unlink(dirname(xml_folder), recursive = TRUE, force = TRUE)
    stop(paste0("Non-English (US) languages detected. Currently, the only ",
                "language encoding supported is 'en-US'."))
  }
}

