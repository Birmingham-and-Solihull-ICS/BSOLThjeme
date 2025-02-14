% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/bsol_presentation_officer.R
\name{bsol_slide_template}
\alias{bsol_slide_template}
\title{Convert to an BSOL ICS Officer Powerpoint presentation}
\usage{
bsol_slide_template(...)
}
\arguments{
\item{...}{additional arguments to pass}
}
\value{
an R script shell that generates a Powerpoint document when run
}
\description{
Format for converting from R to Powerpoint using Birmingham and Solihull ICS branded slide template
}
\examples{
\dontrun{
library(officer)
# this adds slide_template into the environment
bsol_presentation_officer()

my_pres <- read_pptx(slide_template)

# Example slide
my_pres <- add_slide(my_pres, layout = "Title Slide", master="default")

# Write to PowerPoint
print(my_pres, target = "example.pptx")

}
}
