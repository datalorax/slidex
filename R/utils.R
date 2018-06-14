
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

#' Extract xml from pptx
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

#' Import xml code for pptx slides
#'
#' @param xml_folder The folder containing all of the xml code from the pptx,
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

#' Import xml \code{rel} code from pptx
#'
#' @param xml_folder The folder containing all of the xml code from the pptx,
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

extract_class <- function(sld) {
  xml_find_all(sld, "//p:sp/p:nvSpPr/p:nvPr/p:ph") %>%
    map_chr(~xml_attr(., "type"))
}

#' Extract slide title
#'
#' @param sld xml code for the slide to extract the title from
#'
#' @keywords internal

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

# This function is only used within the extract_body function
max_amount <- function(x) {
  x <- na.omit(x)
  if(length(x) == 0) {
    return(0)
  }
  max(x) - 1
}

#' Extract the body of the slide
#'
#' @param sld xml code for the slide to extract the body from
#'
#' @keywords internal

extract_body <- function(sld) {

  sps <- xml_find_all(sld, "//p:sp")
  aps <- map(sps, ~xml_find_all(., "./p:txBody/a:p"))
  classes <- extract_class(sld)

  aps <- aps[-grep("title|ftr", classes)]

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
           bullet  = ifelse(.data$nchar == 0,
                            0,
                            .data$bullet),
           first = ifelse(lag(.data$bullet) == 0 & .data$bullet == 1,
                          1,
                          0),
           first = ifelse(is.na(.data$first),
                          1,
                          .data$first),
           indents = ifelse(.data$first == 1 & .data$bullet == 1,
                            0,
                            .data$indents))

  text <- text %>%
    mutate(set = .data$bullet != lag(.data$bullet),
           set = ifelse(is.na(.data$set), TRUE, .data$set),
           set = cumsum(.data$set)) %>%
    group_by(.data$set) %>%
    mutate(flag = ifelse(.data$indents > lag(.data$indents) + 1,
                         1,
                         0),
           flag = ifelse(is.na(.data$flag), 0, .data$flag),
           flag = cumsum(.data$flag),
           amount = max_amount(.data$indents - lag(.data$indents))*.data$flag,
           indents = ifelse(.data$indents > 1,
                            .data$indents - .data$amount,
                            .data$indents))

  text <- text %>%
    ungroup() %>%
    mutate(spaces = map(.data$indents, ~paste0(paste0(rep("\t", .),
                                                      collapse = ""),
                                               "+")),
           spaces = ifelse(.data$bullet == 0, "", .data$spaces),
           text   = ifelse(.data$bullet != 0,
                           paste(.data$spaces, .data$text),
                           .data$text),
           text   = ifelse(.data$bullet == 0,
                           paste0("\n", .data$text),
                           .data$text)) %>%
    select(.data$text) %>%
    unnest()

  text <- map(text, ~.[. != "\n"])
  paste0(text$text, collapse = "\n")
}

extract_footnote <- function(sld) {
  classes <- extract_class(sld)
  if(!any(grepl("ftr", classes))) {
    return()
  }
  xml_find_all(sld, "//p:sp")[[grep("ftr", classes)]] %>%
    xml_text() %>%
    paste0("\n\n.footnote[", ., "]")
}

# from command line
# libreoffice --headless --convert-to png image.emf

#' Extract attributes from the corresponding slide
#'
#' @param rel xml \code{rel} code for the slide
#' @param attr Attribute to extract. Currently takes two valid arguments:
#'   \code{"image"} or \code{"link"} to extract images or links, respectively.
#' @param sld xml code for the slide to extract the attribute from
#' @keywords internal

# xml_folder will need to be another argument if the commented code below is
# incorporated

extract_attr <- function(rel, attr, sld) {
  types  <- map(xml_children(rel), ~xml_attr(., "Type"))
  target <- map(xml_children(rel), ~xml_attr(., "Target"))

  if(length(target[grep(attr, types)]) == 0) {
    return()
  }
  # if(attr == "chart") {
  #   chart_path <- map_chr(target[grep(attr, types)], ~gsub("\\.\\./", "", .))
  #   chart_path <- map(chart_path, ~c(xml_folder,
  #                                    "ppt",
  #                                    str_split(., "/")[[1]])) %>%
  #     unlist(recursive = FALSE) %>%
  #     reduce(file.path)
  #
  #   chart_xml <- map(chart_path, read_xml)
  #
  #   data <- map(chart_xml, ~xml_find_all(., "//cx:data"))
  #
  #   x_data <- map(data, ~map(.,
  #                          ~xml_find_all(., "./cx:strDim/cx:lvl/cx:pt"))) %>%
  #     map(~map(., xml_text)) %>%
  #     unlist(x_data, recursive = FALSE) %>%
  #     setNames(paste0("V", seq_along(.))) %>%
  #     as.data.frame() %>%
  #     gather(id, x)
  #
  #   y_data <- map(data, ~map(.,
  #                          ~xml_find_all(., "./cx:numDim/cx:lvl/cx:pt"))) %>%
  #      map(~map(., xml_text)) %>%
  #      unlist(recursive = FALSE) %>%
  #      setNames(paste0("V", seq_along(.))) %>%
  #      as.data.frame() %>%
  #      gather(id, y)
  #
  #      full_join(x_data, y_data)
  #
  # }
  if(attr == "link") {
    ar <- xml_find_all(sld, "//a:r")
    select <- xml_find_all(ar, "./a:rPr") %>%
      map(~xml_find_all(., "./a:hlinkClick")) %>%
      map_lgl(~length(.) > 0 )

    links <- target[grep("hyperlink", types)]
    out <- paste0("[", xml_text(ar)[select], "]", "(", links, ")",
                  collapse = "\n")
  }
  out
}

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
      out <- paste0("---\nclass: inverse\nbackground-image: url('assets/img/",
                    imgs, "')\nbackground-size: cover\n",
                    collapse = "\n")
    }
  out
}

#' Extract tables from slides
#'
#' @param sld xml code for the slide to extract the table from
#' @return a \code{data.frame} with the data from the table. Generally fed to
#'   \code{\link{tribble_code}}.
#'
#' @keywords internal

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

#' Wrap a DF in \code{tibble::tribble} code
#'
#' @param df A \code{data.frame}, typically the output from
#'   \code{\link{extract_table}}.
#' @param tbl_num The table number. Not produced in the caption, but used
#'   to name the object and the code chunk. In typical application, corresponds
#'   to the slide number.
#' @keywords internal

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

import_notes_xml <- function(xml_folder) {
  notes_folder <- file.path(xml_folder, "ppt", "notesSlides")
  if(!dir.exists(notes_folder)) {
    return()
  }

  map(list.files(notes_folder, "\\.xml", full.names = TRUE), read_xml)
}

#' Function to extract notes from a slide
#'
#' @param notes A list of the xml code with all the notes for all slides
#' @param sld_num The specific slide number to pull the notes from
#' @param inslides Logical. Should the notes be embedded in the slides?
#'   Defaults to \code{TRUE}.
#'
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

#' Create the \href{https://github.com/yihui/xaringan}{xaringan} YAML front
#' matter
#'
#' @param title_sld The xml code for the title slide in the pptx.
#' @param author A string indicating the author or authors of the slide deck.
#'   Multiple authors can be provided with a string vector.
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
#' @keywords internal

create_yaml <- function(title_sld, author, title = NULL, sub = NULL,
                        date = Sys.Date(), theme = "default",
                        highlightStyle = "github") {
  if(is.null(title)) {
    ttl  <- extract_title(title_sld)
    ttl  <- paste0("title: '",
                   gsub("\t|\n", "", substr(ttl, 3, nchar(ttl))),
                   "'")
  }
  if(is.null(sub)) {
    sub <- extract_subtitle(title_sld)
  }
  if(!is.null(title)) ttl <- paste0("title: '", title, "'")
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

  if(length(author) == 1) {
    auth <- paste0("author: ", author)
  }
  if(length(author) > 1) {
    auth <- paste0("author:\n",
                   paste0("  - ", author, collapse = "\n"))
  }
  elements <- list("---",
                   ttl,
                   sub,
                   auth,
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

sink_rmd <- function(xml_folder, rmd, slds, rels,
                     title_sld, author, title, sub, date, theme,
                     highlightStyle) {

  sld_notes <- import_notes_xml(xml_folder)

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
            extract_image(.x, .y),
            extract_attr(.y, "link", .x),
            extract_footnote(.x),
            extract_notes(sld_notes, .z + 1),
            sep = "\n")
      )
  sink()
}
