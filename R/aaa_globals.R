# global variables are defined in this file. the file has been named such that
# it will be executed before any other file


# global variables are defined in this file. the file has been named such that
# it will be executed before any other file
.onLoad <- function(libname, pkgname){

  sysfonts::font_paths()
  sysfonts::font_add_google("Open Sans", "Open Sans")
  sysfonts::font_add("Arial", "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf")
  sysfonts::font_add("Calibri", "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf")
  sysfonts::font_add("Segoe UI", "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf")
  showtext::showtext_auto()

  #extrafont::font_import()
  # suppressMessages({
  # extrafont::loadfonts(device = "win")
  # })

  #sysfonts::font_add("Arial", "C:\\WINDOWS\\Fonts\\arial.ttf" )
  #showtext::showtext_auto()
}





# BSOL Colours Vector
#
# The colours that are used in the BSOLTheme. Non-exported (used by
# other functions e.g. bsol_theme_cols)
bsol_theme_colours <- c(
  green = "#8cedab",
  light_blue = "#4fbff0",
  orange = "#fc8700",
  deep_navy = "#031d44",
  purple = "#b88ce3",

  nhs_blue = "#005EB8",
  light_slate = "#b2b7b9",
  charcoal = "#2c2825",
  white= "#ffffff"
  #light_charcoal = "#9d928a",
  #light_slate = "#b2b7b9",
  #light_blue = "#abc0e0",
  #light_red = "#f5b2aa",

  #dark_orange = "#7c5f03",
  #dark_charcoal = "#151412",
  #dark_slate = "#343739",
  #dark_blue = "#263f66",
  #dark_red = "#901d10"
)

# BSOL ICS Palettes Vector
#
# The palettes that are used in BSOL ICB. Non-exported (used by
# other functions e.g. bsol_theme_cols)
bsol_theme_palettes <- c(
  "main",
  "alternative",
  "ics_blue",
  "ics_orange",
  "ics_green",
  "ics_purple",
  "ics_navy"
)



# taken from usethis: https://github.com/r-lib/usethis
#' @noRd
#' @import usethis fs
scoped_path_r <- function(scope = c("user", "project"), ..., envvar = NULL) {
  scope <- match.arg(scope)

  # Try environment variable in user scopes
  if (scope == "user" && !is.null(envvar)) {
    env <- Sys.getenv(envvar, unset = "")
    if (!identical(env, "")) {
      return(user_path_prep(env))
    }
  }

  root <- switch(scope,
                 user = fs::path_home_r(),
                 project = usethis::proj_get()
  )
  path(root, ...)
}

user_path_prep <- function(path) {
  ## usethis uses fs's notion of home directory
  ## this ensures we are consistent about that
  path_expand(path)
}

create_directory <- function(path) {
  if (dir_exists(path)) {
    return(invisible(FALSE))
  } else if (file_exists(path)) {
    ui_stop("{ui_path(path)} exists but is not a directory.")
  }

  dir_create(path, recurse = TRUE)
  ui_done("Creating {ui_path(path)}")
  invisible(TRUE)
}
