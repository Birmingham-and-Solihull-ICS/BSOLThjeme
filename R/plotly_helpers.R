
#' ggplotly BSOL wrapper for passing the font size properly
#'
#' @param p a ggplot object, default is the last plot called
#' @param font required font, extracted from the plot if not specified otherwise
#'
#' @inheritParams plotly::layout
#'
#' @return A plotly object
#' @export
#'
#' @examples
#' library(ggplot2)
#' a<-ggplot(iris, aes(y=Sepal.Length, x=Sepal.Width, colour=Species))+
#' geom_point()+
#' scale_colour_bsol()+
#' theme_bsol()
#'
#' ggplotly_bsol(a)
#'
ggplotly_bsol <- function(p = ggplot2::last_plot(), font, ...){

  # check ggplot class input
  if(!any(grepl("ggplot", class(p)))){
    stop("Input plot is not in ggplot format")
  }

  if(missing(font)){
    font <- list(
      p$theme$text$family
      , p$theme$text$size
    )
  } else {
    font <- list(
      font
      , p$theme$text$size
    )
  }

  out <-
    plotly::layout(plotly::ggplotly(p)
                   , font = font
                   , legend = list(title= list(font=font), font = font)
                   , title = list(font = font)
                   , xaxis= list(title = list(font=font), tickfont = font)
                   , yaxis= list(title = list(font=font), tickfont = font)
                   )

  return(out)
}


class
