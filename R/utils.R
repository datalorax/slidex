#' Extract Author from PowerPoint
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' @keywords internal
#'
extract_author <- function(xml_folder) {
  core <- read_xml(file.path(xml_folder, "docProps", "core.xml"))
  xml_find_all(core, "//dc:creator") %>%
    xml_text()
}

#' Extract XML from PowerPoint
#'
#' @param path Path to the Microsoft PowerPoint file
#' @param force If an 'assets' folder already exists in the current directory,
#'   (e.g., from a previous conversion) should it be overwritten? Defaults to
#'   \code{FALSE}.
#' @keywords internal

extract_xml <- function(path, force = FALSE) {
  ppt <- basename(path)
  folder <- gsub("\\.pptx", "", ppt)
  tmpdir <- tempdir()
  if(dir.exists(tmpdir)) {
    unlink(tmpdir, recursive = TRUE, force = TRUE)
  }
  dir.create(tmpdir, showWarnings = FALSE)
  basepath <- file.path(tmpdir, folder)

  dir.create(basepath, showWarnings = FALSE)
  dir.create(file.path(basepath, "xml"), showWarnings = FALSE)

  file.copy(path, file.path(basepath, "xml", ppt))
  file.rename(file.path(basepath, "xml", ppt),
              gsub("\\.pptx", "\\.zip", file.path(basepath, "xml", ppt)))

  unzip(gsub("\\.pptx", "\\.zip", file.path(basepath, "xml", ppt)),
        exdir = file.path(basepath, "xml"))

  if(file.exists(file.path(basepath, "xml", "ppt", "media"))) {
    dir.create(file.path(basepath, "assets"), showWarnings = FALSE)
    dir.create(file.path(basepath, "assets", "img"), showWarnings = FALSE)
    file.rename(file.path(basepath, "xml", "ppt", "media"),
                file.path(basepath, "assets", "img"))
  }
  if(file.exists(file.path(basepath, "xml", "ppt", "embeddings"))) {
    dir.create(file.path(basepath, "assets", "data"), showWarnings = FALSE)
    file.rename(file.path(basepath, "xml", "ppt", "embeddings"),
                file.path(basepath, "assets", "data"))
  }
  rels <- list.files(file.path(basepath, "xml", "ppt", "slides", "_rels"),
                     full.names = TRUE)

  invisible(file.rename(rels, substr(rels, 1, nchar(rels) - 5)))
  invisible(file.path(basepath, "xml"))
}

#' Import XML for PowerPoint Slides
#'
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' created from \code{\link{extract_xml}}.
#' @keywords internal

import_slide_xml <- function(xml_folder) {
  slds <- file.path(xml_folder, "ppt", "slides") %>%
    list.files(pattern = "\\.xml", full.names = TRUE) %>%
    map(read_xml) %>%
    setNames(
      list.files(
        file.path(xml_folder, "ppt", "slides"),
        pattern = "\\.xml"))

  order <- strsplit(names(slds), "/") %>%
    map_chr(~.[length(.)]) %>%
    gsub("[^[:digit:]]", "", .) %>%
    as.numeric() %>%
    order()

  slds[order]
}

#' Import XML \code{rel} Code from PowerPoint
#'
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' created from \code{\link{extract_xml}}.
#' @keywords internal

import_rel_xml <- function(xml_folder) {
  rels <- file.path(xml_folder, "ppt", "slides", "_rels") %>%
    list.files(pattern = "\\.xml", full.names = TRUE) %>%
    map(read_xml) %>%
    setNames(
      list.files(
        file.path(xml_folder, "ppt", "slides"),
        pattern = "\\.xml"))

   order <- strsplit(names(rels), "/") %>%
    map_chr(~.[length(.)]) %>%
    gsub("[^[:digit:]]", "", .) %>%
    as.numeric() %>%
    order()

  rels[order]
}


#' Extract Slide Element Classes
#'
#' @param sld xml code for the slide to extract the title from
#' @keywords internal
#'
extract_class <- function(sld) {
  xml_find_all(sld, "//p:sp/p:nvSpPr/p:nvPr/p:ph") %>%
    map_chr(~xml_attr(., "type"))
}


#' Extract Slide Titles
#'
#' @param sld xml code for the slide to extract the title from
#' @keywords internal
#'

extract_title <- function(sld) {
  classes <- extract_class(sld)

  title <- xml_find_all(sld, "//p:sp/p:txBody")[grep("[tT]itle", classes)] %>%
    xml_text()
  if(length(grep("subTitle", classes)) != 0) {
    title <- title[-grep("subTitle", classes)]
  }

  out <- paste("# ", title, "\n")
  out[!grepl("#   \n", out)]
}

#' Extract Subtitle from Title Slide
#'
#' @param sld xml code for the slide to extract the title from
#' @keywords internal

extract_subtitle <- function(sld) {
  classes <- extract_class(sld)
  if(length(grep("subTitle", classes)) == 0) {
    return()
  }

  out <- xml_find_all(sld, "//p:sp/p:txBody")[grep("subTitle", classes)] %>%
    xml_find_all("./a:p") %>%
    map_chr(xml_text) %>%
    paste0(collapse = " ") %>%
    paste0("'", ., "'")
  if(out == "''") {
    return()
  }
  out
}

#' Extract Footnotes from Slides
#'
#' @param sld xml code for the slide to extract the title from
#' @keywords internal

extract_footnote <- function(sld) {
  classes <- extract_class(sld)
  if(!any(grepl("ftr", classes))) {
    return()
  }
  xml_find_all(sld, "//p:sp")[[grep("ftr", classes)]] %>%
    xml_text() %>%
    paste0("\n\n.footnote[", ., "]")
}

#' Import Slide Notes
#'
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' @keywords internal
#'
import_notes_xml <- function(xml_folder) {
  notes_folder <- file.path(xml_folder, "ppt", "notesSlides")
  if(!dir.exists(notes_folder)) {
    return()
  }

  map(list.files(notes_folder, "\\.xml", full.names = TRUE), read_xml)
}

#' Extract Notes from Slides
#'
#' @param notes A list of the xml code with all the notes for all slides
#'   (e.g., the results of \code{import_notes_xml})
#' @param sld_num Numeric. The specific slide number to pull the notes from
#' @param inslides Logical. Should the notes be embedded in the slides?
#'   Defaults to \code{TRUE}.
#' @keywords internal

extract_notes <- function(notes, sld_num, inslides = TRUE) {

  sld_notes_num <- map_dbl(notes,
                       ~xml_find_all(., "//p:txBody/a:p/a:fld/a:t") %>%
                         xml_text(.) %>%
                         as.numeric())

  if(!(sld_num %in% sld_notes_num)) {
    return()
  }
  note <- notes[sld_num == sld_notes_num][[1]]
  note_text <- xml_find_all(note, "//p:txBody/a:p/a:r") %>%
    xml_text(trim = TRUE) %>%
    paste0(collapse = " ")

  if(inslides) {
    out <- paste0("\n???\n", note_text, "\n", collapse = "")
    return(out[-grep(paste0("\n???\n", " ", "\n", collapse = ""), out)])
  }
  if(!inslides) {
    return(paste0(note_text, "\n", collapse = ""))
  }
}

#' Write Slide Notes
#'
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' @keywords internal
#'
write_notes <- function(xml_folder) {
  notes <- import_notes_xml(xml_folder)
  n_slides <- length(
                     list.files(file.path(xml_folder, "ppt", "slides"),
                                "\\.xml")
                     )
  folder <- map_chr(strsplit(dirname(xml_folder), "/"), ~.[[length(.)]])
  notes_out <- file.path(dirname(xml_folder),
                         paste0(folder, "-notes.txt"))
  sink(notes_out)
    map(seq_len(n_slides),
        ~paste0("\n",
                "---",
                "#", .,
                "\n",
                extract_notes(notes, ., inslides = FALSE),
                collapse = "\n")) %>%
      paste0(collapse = "\n") %>%
      cat()
  sink()
}

#' Write Out the \href{https://github.com/yihui/xaringan}{xaringan} RMD File
#' @inheritParams create_yaml
#' @param xml_folder The folder containing all of the xml code from the pptx
#' @param rmd String. Name of the R Markdown file to be written out.
#' @param slds The xml code for all slides, i.e., the output from
#'   \code{import_slide_xml}.
#' @param rels The rel xml code for all slides, i.e., the output from
#'   \code{import_rel_xml}.
#' @keywords internal

write_rmd <- function(xml_folder, rmd, slds, rels,
                     title_sld, author, title, sub, date, theme,
                     highlightStyle) {

  sld_notes <- import_notes_xml(xml_folder)

  sink(rmd)
    cat(
      create_yaml(xml_folder, title_sld, author, title, sub, date, theme,
                  highlightStyle)
    )
    pmap(list(.x = slds, .y = rels, .z = seq_along(slds)),
        function(.x, .y, .z)
        cat("\n---",
            extract_title(.x),
            extract_body(.x),
            tribble_code(extract_table(.x), tbl_num = .z),
            extract_image(.x, .y),
            extract_link(.x, .y),
            extract_footnote(.x),
            extract_notes(sld_notes, .z + 1),
            sep = "\n")
      )
    on.exit(sink())
}
