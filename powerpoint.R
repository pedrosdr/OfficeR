setwd('C:/Users/Lenovo/Documents/projects/OfficeR')
if(!require(officer)) install.packages('officer')
library(officer)

doc = officer::read_pptx()
doc = officer::add_slide(doc, layout = 'Title and Content', master = "Office Theme")
doc = officer::add_slide(doc, layout = 'Two Content', master = "Office Theme")

# doc = move_slide(doc, index = 1, to = 2)
# doc = remove_slide(doc, index = 1)

doc = officer::on_slide(doc, index = 1)
doc = officer::ph_with(doc, 'My first title', location = ph_location_type(type = 'title'))
doc = officer::ph_with(doc, 'My first paragrath', location = ph_location_type(type = 'body'))
doc = officer::ph_with(doc, 'My first footer', location = ph_location_type(type = 'ftr'))
doc = officer::ph_with(doc, 'Page number 01', location = ph_location_type(type = 'sldNum'))
doc = officer::ph_with(doc, 'Some other text', location = ph_location_type(type = 'dt'))

doc = officer::on_slide(doc, index = 2)
doc = officer:ph_with(doc, "I am going to be left", location = ph_location_left())
doc = officer:ph_with(doc, "I am going to be right", location = ph_location_right())

print(doc, 'document.pptx')
