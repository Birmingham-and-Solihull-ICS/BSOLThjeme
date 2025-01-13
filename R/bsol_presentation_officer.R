#' Convert to an BSOL ICS Officer Powerpoint presentation
#'
#' Format for converting from R to Powerpoint using Birmingham and Solihull ICS branded slide template
#'
#' @param ... additional arguments to pass
#'
#' @return an R script shell that generates a Powerpoint document when run
#'
#' @import officer
#' @importFrom rvg dml
#'
#' @examples
#' \dontrun{
#' library(officer)
#' # this adds slide_template into the environment
#' bsol_presentation_officer()
#'
#' my_pres <- read_pptx(slide_template)
#'
#' # Example slide
#' my_pres <- add_slide(my_pres, layout = "Title Slide", master="default")
#'
#' # Write to PowerPoint
#' print(my_pres, target = "example.pptx")
#'
#' }
#' @export
bsol_slide_template <- function(...) {

  # get the locations of resource files located within the package
  slide_template <- system.file("officer/",
                     "BSOL_DEFAULTw.pptx",
                     package = "BSOLTheme")

 slide_template

}
