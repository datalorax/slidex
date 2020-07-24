
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
