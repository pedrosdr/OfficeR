setwd('C:/Users/Lenovo/Documents/projects/OfficeR')
if(!require(officer)) install.packages('officer')
library(officer)

doc = officer::read_pptx()
doc = officer::add_slide(doc, layout = 'Title and Content', master = "Office Theme")
doc = officer::add_slide(doc, layout = 'Two Content', master = "Office Theme")

doc = move_slide(doc, index = 1, to = 2)

print(doc, 'document.pptx')


