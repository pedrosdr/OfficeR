# layout_summary
setwd('C:/Users/Lenovo/Documents/projects/OfficeR')
if(!require(officer)) install.packages('officer')
library(officer)
library(plotly)
library(ggplot2)

doc = officer::read_pptx()

## Creating slides
# doc = officer::add_slide(doc, layout = 'Title and Content', master = "Office Theme")
# doc = officer::add_slide(doc, layout = 'Two Content', master = "Office Theme")

## Slides basics
# doc = move_slide(doc, index = 1, to = 2)
# doc = remove_slide(doc, index = 1)

## Text
# doc = officer::on_slide(doc, index = 1)
# doc = officer::ph_with(doc, 'My first title', location = ph_location_type(type = 'title'))
# doc = officer::ph_with(doc, 'My first paragrath', location = ph_location_type(type = 'body'))
# doc = officer::ph_with(doc, 'My first footer', location = ph_location_type(type = 'ftr'))
# doc = officer::ph_with(doc, 'Page number 01', location = ph_location_type(type = 'sldNum'))
# doc = officer::ph_with(doc, 'Some other text', location = ph_location_type(type = 'dt'))
# 
# doc = officer::on_slide(doc, index = 2)
# doc = officer::ph_with(doc, "I am going to be left", location = ph_location_left())
# doc = officer::ph_with(doc, "I am going to be right", location = ph_location_right())

# # Tables and Locations
# emp = data.frame(
#   name = c('john', 'jenny', 'margareth'),
#   salary = c(2000, 4300, 5000)
# )
# 
# doc = add_slide(doc, layout = 'Title and Content', master = 'Office Theme')
# doc = on_slide(doc, index = 1)
# 
# doc = ph_with(doc, emp, location = ph_location(
#   left = 0.5,
#   top = 0.5,
#   width = 4,
#   height = 3.5,
#   rotation = 0,
#   bg = "#333333"
# ))

# # Text Formatting
# paragraph = fpar(
#   ftext("Here is the first paragraph...", 
#         fp_text(color = '#48a3d1', font.size = 25)
#   ),
#   ftext('so i am the second part of this paragraph',
#         fp_text(color = 'white', font.size = 12, bold = TRUE)
#   )
# )
# 
# loc = ph_location(
#   left = 2,
#   top = 1,
#   width = 3,
#   height = 2,
#   rotation = 30,
#   bg = "#7905a5"
# )
# 
# doc = ph_with(doc, paragraph, loc)

doc = officer::read_pptx('template.pptx')
doc = on_slide(doc, index = 1)
doc = ph_with(doc, "Hello man", location = ph_location_type(type = 'title'))

doc 

# doc = add_slide(doc, layout = 'Slide de Título', 'Tema do Office')

print(doc, 'document.pptx')
