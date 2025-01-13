library(sysfonts)
library(tidyverse)
library(officer)
library(BSOLTheme)

# Load from the default slide
bsol_presentation_officer()
my_pres <- read_pptx(slide_template)

# Set ggplot theme general
# Set ggplot theme general
theme_set(
  theme_minimal()+
    theme(
      text = element_text(size = 16)
      , axis.text.x=element_blank()
      , plot.subtitle = element_text(face = "italic", size = 20)
      , axis.text = element_text(size = 18)
      , axis.title = element_text(size = 18)
      , legend.text = element_text(size = 20)
      , legend.title = element_text(size = 22)
    )
)

# Work out what slides and layouts you have to work with
layout_summary(my_pres)

# Add new slide, based on 'Title Slide' layout of 'default' slide master
my_pres <- add_slide(my_pres, layout = "Title Slide", master="default")

# Add title, into the 'Presentation title' placeholder
my_pres <- ph_with(my_pres, value = "This is my title", location =  ph_location_label("Presentation Title"))

# Add a new content slide
my_pres <-add_slide(my_pres, layout = "Normal Slide Picture", master="default")

# Add content, turning a ggplot graph into dml format for better graphics
plot1<-
  ggplot(iris, aes(y=Petal.Length, x=Petal.Width, colour = Species))+
  geom_point()+
  scale_colour_viridis_d() +
   theme_grey(base_family = "Arial")

#ggsave(filename = tempfile("plot1.png"), device = "png")
plot1<-  rvg::dml(ggobj = plot1)

print(plot1)

#rea
  my_pres <- ph_with(my_pres, value = plot1, location =  ph_location_label("Picture"))


# Also add some text into title and commentary boxes

my_pres <- ph_with(my_pres, value = "Iris dataset example", location =  ph_location_label("Slide Title"))

commentary <- paste("Here is some commentary, plotting",
                    nrow(iris),
                    "rows in the Iris dataset"
                    )
my_pres <- ph_with(my_pres, value = commentary, location =  ph_location_label("textbox"))


print(my_pres, target = "example.pptx")
