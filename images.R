library(officer)
# library(flextable)
# library(gt)

doc = read_pptx('template.pptx')
doc = add_slide(doc, layout = 'Em branco', master = 'Tema do Office')
doc = add_slide(doc, layout = 'Em branco', master = 'Tema do Office')
doc = add_slide(doc, layout = 'Em branco', master = 'Tema do Office')

sld_width = slide_size(doc)$width
sld_height = slide_size(doc)$height

# img
doc = on_slide(doc, index = 2)
img = external_img('image.jpeg', width = 10, height = 10)
doc = ph_with(doc, value = img, location = ph_location((sld_width - 6)/2, 1, 6, 6))

# table
doc = on_slide(doc, index = 3)
tbl = data.frame(
  'Nome' = c('Adelier Store', 'Biribi F.C.', 'Bar Do Martinez', 'Jangada Pizza Bar'),
  'QS5' = c(2.31, 1.77, 1.56, 3.22),
  'QS3' = c(234, 211, 232, 121)
)

doc = ph_with(doc, tbl, location = ph_location((sld_width - 8)/2, 1, 8, 2))

# table img
doc = on_slide(doc, index = 4)
img = external_img('tables.png')
doc = ph_with(doc, img, location = ph_location(10, 0.5, 2.1, 6))

print(doc, 'document.pptx')
