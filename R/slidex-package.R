#' slidex: convert Microsoft PowerPoint slides to R Markdown
#'
#' This package is designed to extract information from Microsoft
#' PowerPoint slides, and then put that information into an R Markdown
#' document with a \href{xaringan}{https://github.com/yihui/xaringan} YAML. If
#' the xaringan package is also installed, beautiful html slides can then be
#' produced by knitting the RMD. At present, the package exports one function,
#' `convert_pptx`, which converts a .pptx file to R Markdown. The package is
#' not intended to be all encompassing or provide perfect conversion. Rather,
#' it should get you about 90% of the way there for about 80% of use cases.
#' Given the idiosyncrasies of different slide show presentations, it is
#' expected that some manual editing of the resulting RMD will need to be
#' completed to get the HTML slides to look exactly as you want them, but most
#' of the hard work should be done for you. Importantly, the package maintains
#' and provides the proper code for any images that were embedded in the .pptx
#' slides, as well as links. Tables are also generally maintained, although
#' they may require some manual editing if complex spanner heads and merged
#' cells were used in the original table. Nested bulleted lists should also be
#' maintained.
#'
#' @keywords internal
#' @importFrom xml2 read_xml xml_find_all xml_attr xml_text xml_children
#' @importFrom purrr map map2 map_dbl map_chr map_df pmap map_lgl
#' @importFrom magrittr %>%
#' @importFrom dplyr mutate bind_cols select group_by ungroup lag
#' @importFrom tibble tibble
#' @importFrom tidyr unnest
#' @importFrom rlang .data
#' @importFrom utils unzip capture.output write.csv
#' @importFrom stats setNames na.omit
"_PACKAGE"

if(getRversion() >= "2.15.1")  utils::globalVariables(c("."))
