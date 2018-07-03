
xml_tibble <- function(sld) {
  sps <- xml_find_all(sld, "//p:sp")
  nvpr_name <- map(sps, ~xml_find_all(., "./p:nvSpPr/p:cNvPr")) %>%
    map(., ~xml_attr(., "name"))
  aps <- map(sps, ~xml_find_all(., "./p:txBody/a:p"))
  classes <- extract_class(sld)

  aps <- aps[-grep("title|ftr", classes)]
  nvpr_name <- nvpr_name[-grep("title|ftr", classes)]

  if(length(aps) == 0) {
    return()
  }

  tbl <- tibble(sp = unlist(map2(seq_along(aps),
                               map_dbl(aps, length),
                               ~rep(.x, .y))),
                ap = unlist(map(aps, seq_along)),
                type = unlist(map2(nvpr_name, aps,
                            ~rep(.x, length(.y))))) %>%
    bind_cols(tibble(xml = unlist(aps, recursive = FALSE)))

  tbl$type <- gsub("[^content].+", "", tolower(tbl$type))
  tbl$type <- ifelse(tbl$type != "content", "other", tbl$type)

  tbl
}

extract_text <- function(xml_tibble) {
  if(is.null(xml_tibble)) return()
  xml_tibble %>%
    mutate(text  = map(.data$xml, ~xml_find_all(., "./a:r")),
           text  = map(.data$text, ~map(., ~xml_text(., trim = TRUE))))
}

#' bold or italicize text
#' @param text data frame output from extract_text
#' @keywords internal
#'
stylize_text <- function(text) {
  if(is.null(text)) return()
  style <- text %>%
    mutate(style = map(.data$xml, ~xml_find_all(., "./a:r/a:rPr")),
           bold  = map(.data$style, ~ifelse(is.na(xml_attr(., "b")),
                                    "0",
                                    xml_attr(., "b"))),
           ital  = map(.data$style, ~ifelse(is.na(xml_attr(., "i")),
                                    "0",
                                    xml_attr(., "i"))))

  style %>%
    mutate(text = map2(.data$text, .data$bold,
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
}

#' Check bullets before text
#' @param xml_tibble output from extract_text or stylyize_text. Typically the latter.
#' @keywords internal

extract_indents <- function(xml_tibble) {
  if(is.null(xml_tibble)) return()
  xml_tibble %>%
    mutate(indents = map(.data$xml, ~xml_find_all(., "./a:pPr")),
           indents = map(.data$indents, ~xml_attr(., "lvl")),
           indents = map(.data$indents, ~ifelse(is.na(.), "0", .)),
           indents = as.numeric(
             map_chr(.data$indents, ~ifelse(length(.) > 0, ., "0"))),
           indents = ifelse(.data$type == "content",
                            .data$indents + 1,
                            .data$indents),
           nobullet = map_dbl(.data$xml,
                              ~length(xml_find_all(., "./a:pPr/a:buNone"))),
           indents = ifelse(.data$nobullet == 1, 0, .data$indents))
}

max_amount <- function(x) {
  x <- na.omit(x)
  if(length(x) == 0) {
    return(0)
  }
  max(x) - 1
}

#' Fixes the bulleting levels in the case of non-standard nesting
#' @param bulleted A dataframe output from \code{insert_bullets}
#' @keywords internal
#'
level_indents <- function(indented) {
  if(is.null(indented)) return()

  indented %>%
    mutate(set = as.numeric(.data$indents > 0),
           set = .data$set != lag(.data$set),
           set = ifelse(is.na(.data$set), TRUE, .data$set),
           set = cumsum(.data$set)) %>%
    group_by(.data$set) %>%
    mutate(first = ifelse(is.na(lag(.data$indents)), 1, 0),
           indents = ifelse(.data$first == 1 & .data$type == "content",
                            0,
                            .data$indents),
           flag = ifelse(.data$indents > lag(.data$indents) + 1,
                         1,
                         0),
           flag = ifelse(is.na(.data$flag), 0, .data$flag),
           flag = cumsum(.data$flag),
           amount = max_amount(.data$indents - lag(.data$indents))*.data$flag,
           indents = ifelse(.data$indents > 1,
                            .data$indents - .data$amount,
                            .data$indents),
           bullet = (.data$nobullet - 1) * - 1,
           bullet = ifelse(.data$type != "content", 0, .data$bullet)) %>%
    ungroup() %>%
    select(-.data$nobullet, -.data$set, -.data$first, -.data$flag,
           -.data$amount)
}

insert_bullets <- function(indented) {
  if(is.null(indented)) return()

  indented %>%
    mutate(spaces = map(.data$indents, ~paste0(paste0(rep("\t", .),
                                                      collapse = ""),
                                               "+ ")),
           spaces = ifelse(.data$bullet == 0, "", .data$spaces),
           nchar  = map_chr(.data$text, nchar),
           text   = paste0(.data$spaces, .data$text)) %>%
    subset(nchar != "0")
}

#' Text to paste
#' @param text the text to paste into a string
#' @keywords internal
body_text <- function(text) {
  if(is.null(text)) return()
  pasted <- text %>%
    select(.data$text) %>%
    unnest()
  pasted <- map(pasted, ~.[. != "\n"])

  paste0(pasted$text, collapse = "\n")
}

#' Extract the body of the slide
#'
#' @param sld xml code for the slide to extract the body from
#'
#' @keywords internal
extract_body <- function(sld) {
  xml_tibble(sld) %>%
    extract_text() %>%
    stylize_text() %>%
    extract_indents() %>%
    level_indents() %>%
    insert_bullets() %>%
    body_text()
}

