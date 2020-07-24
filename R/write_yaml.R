
#' Write the YAML Code for CSS Theme
#'
#' @param theme The name of the theme, passed as a single string. Note that
#'  just the name needs to be supplied and any accompanying fonts will also
#'  be included.
#'
write_theme <- function(theme) {
  css_dir <- system.file('rmarkdown', 'templates', 'xaringan', 'resources',
                       package = 'xaringan')
  all_themes <- list.files(css_dir, '[.]css$')

  chosen_theme <- all_themes[grepl(theme, all_themes)]
  chosen_theme <- gsub("\\.css", "", chosen_theme)

  paste0('css: ["default", ',
         paste0('"', chosen_theme, '"', collapse = ", "),
         "]",
         collapse = "")
}

#' Create the \href{https://github.com/yihui/xaringan}{xaringan} YAML Front
#' Matter
#'
#' @param xml_folder The folder containing all of the xml code from the pptx.
#' @param title_sld The xml code for the title slide in the pptx.
#' @param author Optional string indicating the author or authors of the slide.
#'   Defaults to the listed creator of the pptx.Multiple authors can be
#'   provided with a string vector.
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

create_yaml <- function(xml_folder, title_sld, author = NULL, title = NULL,
                        sub = NULL, date = Sys.Date(), theme = "default",
                        highlightStyle = "github") {
  if(is.null(author)) {
    author <- extract_author(xml_folder)
  }
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
    css  <- write_theme(theme)
  } else {
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

