#' Set up your environmental variables for BSOL templates
#'
#' Once of the safest, and most reusable ways to automatically add your details to the Rmarkdown templates
#' in this file is to add them to conserved 'environmental variables, that all users share.
#' These variables are MY_NAME, MY_EMAIL, and MY_JOBTITLE. This function scans your .Renviron file (that holds
#' these variables) and detects if they are present.  If not it adds them, then opens the file for
#' you to fill them in.  Once you've done that, please save and restart R, as RStudio caches the file
#' when it opens.
#'
#' @return Updates you user .Renviron file to have the right variables
#' @importFrom usethis edit_file ui_todo ui_done
#' @export
#'
#'
#' @examples
#' \dontrun{
#' bsol_setup()
#' }
bsol_setup <- function(){

  # Environmental variables
  path <- scoped_path_r("user", ".Renviron", envvar = "R_ENVIRON_USER")
  create_directory(path_dir(path))
  file_create(path)

  if(sum(grepl(pattern = "MY_NAME=|MY_EMAIL=|MY_JOBTITLE=|SANDPIT="
               , x = readLines(path))) <3){

    if(sum(grepl(pattern = "MY_NAME=", x = readLines(path)))==0){
      cat("MY_NAME=",file=path, append=TRUE, sep = "\n")
    }
    if(sum(grepl(pattern = "MY_EMAIL=", x = readLines(path)))==0){
      cat("MY_EMAIL=", file=path, append=TRUE, sep = "\n")
    }
    if(sum(grepl(pattern = "MY_JOBTITLE=", x = readLines(path)))==0){
      cat("MY_JOBTITLE=", file=path, append=TRUE, sep = "\n")
    }
    usethis::edit_file(path)
    usethis::ui_todo("Fill in your details in the .Renviron file (no speach-marks/unquoted). Save then restart R for changes to take effect")
    invisible(path)
  }

  usethis::ui_done("You're all set up!")

}
