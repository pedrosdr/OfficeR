library(officer)
library(ggplot2)

doc = read_pptx('template.pptx')
doc = add_slide(doc, layout = 'Em branco', master = 'Tema do Office')
doc = on_slide(doc, index = 2)

emp = data.frame(
  names = c('John', 'Jimmy', 'Yi', 'Mickey'),
  salary = c(200, 220, 180, 300)
)

gg1 = ggplot(emp) +
  geom_col(aes(x = names, y = salary))

size = slide_size(doc)
loc = ph_location(0.5, 0.5, size$width - 1, size$height - 1)

doc = ph_with(doc, gg1, loc)

print(doc, 'document.pptx')
