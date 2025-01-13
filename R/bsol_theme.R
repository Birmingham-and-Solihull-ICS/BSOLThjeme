#' Creates the BSOL Theme for ggplot
#'
#' Creates a base theme for ggplot objects.
#'
#' @param base_family The font to use in this theme
#'
#' @export
#'
#' @import ggplot2 sysfonts showtext
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#' library(BSOLTheme)
#' ggplot(iris, aes(Sepal.Width, Sepal.Length, colour = Species)) +
#'   geom_point() +
#'   theme_bsol()
#' }
theme_bsol <- function(base_family = "Open Sans") {
  theme_classic(base_family = base_family) +
    theme(
      panel.background = element_blank()
      , panel.grid = element_blank()
      , plot.title = element_text(face="bold", size = 16)
      , plot.subtitle = element_text(face="italic",size = 10)
      , axis.line = element_line(colour = "#808080", linewidth = 0.75)
      , strip.background = element_rect(fill = bsol_theme_cols("light_slate"))
      #, panel.grid.major = element_line(colour = "#b3b3b3", linewidth = 0.4)
      #, panel.grid.minor = element_line(colour = "#cfcfcf", linewidth = 0.4)

      )
}


#' Creates the BSOL Theme for ggplot, with Officer package
#'
#' Creates a base theme for ggplot objects, but some of the sizes needed boosting to work with exporting
#' to Powerpoint using Officer.  When translating over, the size gets a bit lost.  This works with the
#' template included in this package.
#'
#' @param base_family The font to use in this theme
#'
#' @export
#'
#' @import ggplot2
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#' library(BSOLTheme)
#' ggplot(iris, aes(Sepal.Width, Sepal.Length, colour = Species)) +
#'   geom_point() +
#'   theme_bsol_officer()
#' }
theme_bsol_officer <- function(base_family = "Open Sans") {
  theme_classic(base_family = base_family) +
    theme(
      panel.background = element_blank()
      , panel.grid = element_blank()
      , plot.title = element_text(face="bold", size = 22)
      #, plot.subtitle = element_text(face="italic",size = 10)
      , axis.line = element_line(colour = "#808080", linewidth = 0.75)
      , strip.background = element_rect(fill = BSOLTheme::bsol_theme_cols("light_slate"))
      , text = element_text(size = 16)
      , axis.text.x=element_blank()
      , plot.subtitle = element_text(face = "italic", size = 20)
      , axis.text = element_text(size = 18)
      , axis.title = element_text(size = 18)
      , legend.text = element_text(size = 20)
      , legend.title = element_text(size = 22)

    )
}

