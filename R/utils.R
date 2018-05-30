#' Extract xml from pptx
#'
#' @param path Path to the Microsoft PowerPoint file

extract_xml <- function(path) {
  ppt_splt <- strsplit(path, "/")
  ppt <- map_chr(ppt_splt, ~.[length(.)])

  xml_folder <- paste0(gsub("\\.pptx", "", ppt), "_xml")
  dir.create(xml_folder, showWarnings = FALSE)
  dir.create("assets", showWarnings = FALSE)

  file.copy(path, file.path(xml_folder, ppt))
  file.rename(file.path(xml_folder, ppt),
              gsub("\\.pptx", "\\.zip", file.path(xml_folder, ppt)))

  unzip(gsub("\\.pptx", "\\.zip", file.path(xml_folder, ppt)),
        exdir = xml_folder)

  if(file.exists(file.path(xml_folder, "ppt", "media"))) {
    dir.create(file.path("assets", "img"), showWarnings = FALSE)
    file.rename(file.path(xml_folder, "ppt", "media"),
                file.path("assets", "img"))
  }
  if(file.exists(file.path(xml_folder, "ppt", "embeddings"))) {
    dir.create(file.path("assets", "data"), showWarnings = FALSE)
    file.rename(file.path(xml_folder, "ppt", "embeddings"),
                file.path("assets", "data"))
  }
  rels <- list.files(file.path(xml_folder, "ppt", "slides", "_rels"),
                     full.names = TRUE)

  invisible(file.rename(rels, substr(rels, 1, nchar(rels) - 5)))
  invisible(xml_folder)
}

#' Import xml Code for PPTX Slides
#'
#' @param xml_folder The folder containing all of the xml code from the pptx,
#' created from \code{\link{extract_xml}}.

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

#' Import xml \code{rel} Code from PPTX
#'
#' @param xml_folder The folder containing all of the xml code from the pptx,
#' created from \code{\link{extract_xml}}.

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


#' Extract Slide Title
#'
#' @param sld xml code for the slide to extract the title from
#'
extract_title <- function(sld) {
  classes <- xml_find_all(sld, "//p:sp/p:nvSpPr/p:cNvPr") %>%
    xml_attr("name")

  title <- xml_find_all(sld, "//p:sp/p:txBody")[grep("Title", classes)] %>%
    xml_text()
  paste("# ", title, "\n")
}

#' Extract the body of the slide
#'
#' @param sld xml code for the slide to extract the title from
#'
extract_body <- function(sld) {

  sps <- xml_find_all(sld, "//p:sp")
  aps <- map(sps, ~xml_find_all(., "./p:txBody/a:p"))
  classes <- xml_find_all(sld, "//p:sp/p:nvSpPr/p:cNvPr") %>%
    xml_attr("name")

  aps <- aps[-grep("Title|Slide Number", classes)]

  if(length(aps) == 0) {
    return()
  }

  d <- tibble(sps = unlist(map2(seq_along(aps),
                               map_dbl(aps, length),
                               ~rep(.x, .y))),
             aps = unlist(map(aps, seq_along))) %>%
    bind_cols(tibble(xml = unlist(aps, recursive = FALSE)))

  text <- d %>%
    mutate(text  = map(.data$xml, ~xml_find_all(., "./a:r")),
           text  = map(.data$text, ~map(., ~xml_text(., trim = TRUE))),
           style = map(.data$xml, ~xml_find_all(., "./a:r/a:rPr")),
           bold  = map(.data$style, ~ifelse(is.na(xml_attr(., "b")),
                                    "0",
                                    xml_attr(., "b"))),
           ital  = map(.data$style, ~ifelse(is.na(xml_attr(., "i")),
                                    "0",
                                    xml_attr(., "i"))),
           text = map2(.data$text, .data$bold,
                       ~map2(.x, .y,
                             ~ifelse(.y == "1",
                                     paste0("**", .x, "**"),
                                     .x))),
           text = map2(.data$text, .data$ital,
                       ~map2(.x, .y,
                             ~ifelse(.y == "1",
                                     paste0("*", .x, "*"),
                                     .x))),
           text = map(.data$text, ~paste(., collapse = " "))) %>%
    select(-.data$style, -.data$bold, -.data$ital)

  text <- text %>%
    mutate(nchar   = map_dbl(.data$text, nchar),
           indents = map(.data$xml, ~xml_find_all(., "./a:pPr")),
           indents = map(.data$indents, ~xml_attr(., "lvl")),
           indents = map(.data$indents, ~ifelse(is.na(.), "0", .)),
           indents = as.numeric(
             map_chr(.data$indents, ~ifelse(length(.) > 0, ., "0"))),
           bullet  = map(.data$xml, ~xml_find_all(., "./a:pPr/a:buNone")),
           bullet  = -1*(map_dbl(.data$bullet, length) - 1),
           bullet  = ifelse(.data$nchar == 0, 0, .data$bullet),
           spaces = map(.data$indents, ~paste0(paste0(rep("\t", .),
                                                collapse = ""), "+")),
           # spaces = ifelse(.data$indents > 2,
           #                 paste0("\t", .data$spaces),
           #                 .data$spaces),
           spaces = ifelse(.data$bullet == 0, "", .data$spaces),
           text   = ifelse(.data$bullet != 0, paste(.data$spaces, .data$text), .data$text),
           text   = ifelse(.data$bullet == 0, paste0("\n", .data$text), .data$text)) %>%
    select(.data$text) %>%
    unnest()

  text <- map(text, ~.[. != "\n"])
  paste0(text$text, collapse = "\n")
}

# from command line
# libreoffice --headless --convert-to png image.emf

#' Extract Attributes from the corresponding slide
#'
#' @param rel xml \code{rel} code for the slide
#' @param attr Attribute to extract. Currently takes two valide arguments:
#'   \code{"image"} or \code{"link"} to extract images or links, respectively.
#' @param sld xml code for the slide to extract the title from

extract_attr <- function(rel, attr, sld) {
  types  <- map(xml_children(rel), ~xml_attr(., "Type"))
  target <- map(xml_children(rel), ~xml_attr(., "Target"))

  if(length(target[grep(attr, types)]) == 0) {
    return()
  }

  if(attr == "link") {
    ar <- xml_find_all(sld, "//a:r")
    select <- xml_find_all(ar, "./a:rPr") %>%
      map(~xml_find_all(., "./a:hlinkClick")) %>%
      map_lgl(~length(.) > 0 )

    links <- target[grep("hyperlink", types)]
    out <- paste0("[", xml_text(ar)[select], "]", "(", links, ")", collapse = "\n")
  }
  if(attr == "image") {
    imgs <- target[grep("image", types)]
    splt <- map(imgs, strsplit, "/")
    imgs <- map_chr(splt, ~map_chr(., ~.[[length(.)]]))

    out <- paste0("![](assets/img/", imgs, ")")
    if(length(out) == 2) {
      out <- paste0(".pull-left[", out[1], "]", "\n\n",
                    ".pull-right[", out[2], "]")
    }
    if(length(out) > 2) {
      out <- paste0("---\nclass: inverse\nbackground-image: url('assets/img/",
                    imgs, "')\nbackground-size: cover\n",
                    collapse = "\n")
    }
  }
  out
}

#' Extract tables from slides
#'
#' @param sld xml code for the slide to extract the title from
#' @return a \code{data.frame} with the data from the table. Generally fed to
#'   \code{\link{tribble_code}}.
#'

extract_table <- function(sld) {
  rows  <- xml_find_all(sld, "//a:tr")
  if(length(rows) == 0) {
    return()
  }
  cols <- map(rows, ~xml_find_all(., "./a:tc"))
  ar <- map(cols, ~map(., ~xml_find_all(., "./a:txBody/a:p/a:r")))

  txt <- map(ar, ~map(., ~map(., ~xml_text(., trim = TRUE))))
  txt <- map(txt, ~map(., paste0, collapse = " "))
  txt <- map(txt, ~map(., ~ifelse(nchar(.) == 0, " ", .)))

  df <- map_df(txt, ~as.data.frame(t(unlist(.)), stringsAsFactors = FALSE))

  names(df) <- df[1, ]
  df <- df[-1, ]
  df
}

#' Wrap a DF in \code{tibble::tribble} Code
#'
#' @param df A \code{data.frame}, typically the output from
#'   \code{\link{extract_table}}.
#' @param tbl_num The table number. Not produced in the caption, but used
#'   to name the object and the code chunk. In typical application, corresponds
#'   to the slide number.


tribble_code <- function(df, tbl_num = "") {

  if(is.null(df)) return()

  nms <- ifelse(nchar(names(df)) == 0, ".", names(df))

  out <- capture.output(write.csv(df, quote = TRUE, row.names = FALSE))
  out <- paste0(c(paste0("~`", nms, "`", collapse = ", "),out[-1]))

  paste0(
    "\n```{r ", paste0("tbl", tbl_num), ", echo = FALSE}\n",
    paste0("tbl", tbl_num, " <- tibble::tribble(\n",
           paste(out, collapse = ",\n"), "\n)"),
    "\n\n",
    paste0("kableExtra::kable_styling(knitr::kable(tbl", tbl_num, "), ",
                                      "font_size = 18)"),
    "\n```"
  )
}

#' Create the \href{https://github.com/yihui/xaringan}{xaringan} YAML Front
#' Matter
#'
#' @param title_sld The xml code for the title slide in the PPTX.
#' @param author The author of the slide deck. Currently a required argument.
#' @param title Optional title of the slide deck. Defaults to the title of the
#'   first slide in the deck.
#' @param sub Optional subtitle
#' @param date The date the slides were produced. Defaults to current date.
#' @param theme The css theme to apply to the xaringan slides. For options, see
#' \href{https://github.com/yihui/xaringan/tree/master/inst/rmarkdown/templates/xaringan/resources}{here}.
#' Note that only the name of the theme needs to be applied
#' (e.g., \code{theme = "metropolis"}) and the proper code will be applied to
#' load both the theme and the fonts, although this can easily be manually
#' manipulated after conversion if you want other fonts with a specific theme.
#' @param highlightStyle The code highlighting style. Defaults to
#'   \code{"github"} flavored highlighting

create_yaml <- function(title_sld, author, title = NULL, sub = NULL,
                        date = Sys.Date(), theme = "default",
                        highlightStyle = "github") {
  if(is.null(title)) {
    ttl  <- extract_title(title_sld)
    ttl  <- paste0("title: ", gsub("\t|\n", "", substr(ttl, 3, nchar(ttl))))
  }
  if(!is.null(title)) ttl <- paste0("title: ", title)
  if(!is.null(sub)) sub  <- paste0("subtitle: ", sub)

  date <- paste0("date: ", date)
  hls  <- paste0("highlightStyle: ",  highlightStyle)

  if(theme != "default") {
    css  <- paste0('    css: ["default", ',
                   '"', theme, '", ',
                   '"', paste0(theme, '-fonts"]'))
  }
  else{
    css <- NULL
  }

  elements <- list("---",
                   ttl,
                   sub,
                   paste0("author: ", author),
                   date,
                   "output:",
                   "  xaringan::moon_reader:",
                   css,
                   "  lib_dir: libs",
                   "  nature:",
                   paste0("    ", hls),
                   "    highlightLines: true",
                   "    countIncrementalSlides: false")
  elements <- elements[!map_lgl(elements, is.null)]

  paste0(elements, collapse = "\n")
}
