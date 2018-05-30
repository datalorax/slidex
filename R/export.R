#' Extract xml from pptx
#'
#' @param path Path to the Microsoft PowerPoint file
#' @inheritParams create_yaml
#' @export
#'
convert_pptx <- function(path, author, title = NULL, sub = NULL,
                         date = Sys.Date(), theme = "default",
                         highlightStyle = "github") {
  xml  <- extract_xml(path)
  slds <- import_slide_xml(xml)
  rels <- import_rel_xml(xml)

  title_sld <- slds[[1]]

  slds <- slds[-1]
  rels <- rels[-1]

  rmd <- paste0(gsub("_xml", "", xml), ".Rmd")

  sink(rmd)
    cat(
      create_yaml(title_sld, author, title, sub, date, theme, highlightStyle)
    )
    pmap(list(.x = slds, .y = rels, .z = seq_along(slds)),
        function(.x, .y, .z)
        cat("\n---",
            extract_title(.x),
            extract_body(.x),
            tribble_code(extract_table(.x), tbl_num = .z),
            extract_attr(.y, "image", .x),
            extract_attr(.y, "link", .x),
            sep = "\n")
      )
  sink()
  unlink("slidedemo_xml", recursive = TRUE)
  system(paste("open", rmd))
}


