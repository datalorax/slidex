extract_text <- function(sld) {
  sps <- xml_find_all(sld, "//p:sp")
  aps <- map(sps, ~xml_find_all(., "./p:txBody/a:p"))
  classes <- extract_class(sld)

  aps <- aps[-grep("title|ftr", classes)]

  if(length(aps) == 0) {
    return()
  }

    tibble(sps = unlist(map2(seq_along(aps),
                               map_dbl(aps, length),
                               ~rep(.x, .y))),
             aps = unlist(map(aps, seq_along))) %>%
    bind_cols(tibble(xml = unlist(aps, recursive = FALSE))) %>%
    mutate(text  = map(.data$xml, ~xml_find_all(., "./a:r")),
           text  = map(.data$text, ~map(., ~xml_text(., trim = TRUE))))
}

#' bold or italicize text
#' @param text data frame output from extract_text
#' @keywords internal
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
  stylized <- style %>%
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

  stylized
}

#' Insert bullets before text
#' @param text output from extract_text or stylyize_text. Typically the latter.
#' @keywords internal
check_bullets <- function(text) {
  if(is.null(text)) return()
  indents <- text %>%
    mutate(nchar   = map_dbl(.data$text, nchar),
           indents = map(.data$xml, ~xml_find_all(., "./a:pPr")),
           indents = map(.data$indents, ~xml_attr(., "lvl")),
           indents = map(.data$indents, ~ifelse(is.na(.), "0", .)),
           indents = as.numeric(
             map_chr(.data$indents, ~ifelse(length(.) > 0, ., "0"))))

  indents %>%
    mutate(bullet  = map(.data$xml, ~xml_find_all(., "./a:pPr/a:buNone")),
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
level_bullets <- function(bulleted) {
  if(is.null(bulleted)) return()
  bulleted%>%
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
}

insert_bullets <- function(leveled_bullets) {
  if(is.null(leveled_bullets)) return()
  leveled_bullets %>%
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
                           .data$text))
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
  extract_text(sld) %>%
    stylize_text() %>%
    check_bullets() %>%
    level_bullets() %>%
    insert_bullets() %>%
    body_text()
}
