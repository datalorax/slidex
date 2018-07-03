#' Extract Links from the corresponding slide
#'
#' @param sld xml code for the slide to extract the attribute from
#' @param rel xml \code{rel} code for the slide
#' @keywords internal

extract_link <- function(sld, rel) {
  types  <- map(xml_children(rel), ~xml_attr(., "Type"))
  target <- map(xml_children(rel), ~xml_attr(., "Target"))

  if(length(target[grep("hyperlink", types)]) == 0) {
    return()
  }

  ar <- xml_find_all(sld, "//a:r")
  select <- xml_find_all(ar, "./a:rPr") %>%
    map(~xml_find_all(., "./a:hlinkClick")) %>%
    map_lgl(~length(.) > 0 )

  links <- target[grep("hyperlink", types)]

  paste0("[", xml_text(ar)[select], "]", "(", links, ")",
         collapse = "\n")

}

# from command line
# libreoffice --headless --convert-to png image.emf

extract_image <- function(sld, rel) {
  sld_ids <- xml_find_all(sld, "//p:pic/p:blipFill/a:blip") %>%
    xml_attr(., "embed")
  if(length(sld_ids) == 0) {
    return()
  }
  rel_ids <- map_chr(xml_children(rel), ~xml_attr(., "Id"))
  imgs <- map(xml_children(rel), ~xml_attr(., "Target")) %>%
    .[rel_ids %in% sld_ids] %>%
    map_chr(basename)

    out <- paste0("![](assets/img/", imgs, ")")
    if(length(out) == 2) {
      out <- paste0(".pull-left[", out[1], "]", "\n\n",
                    ".pull-right[", out[2], "]")
    }
    if(length(out) > 2) {
      out <- paste0("\n---\nclass: inverse\nbackground-image: url('assets/img/",
                    imgs, "')\nbackground-size: cover\n",
                    collapse = "\n")
    }
  out
}
