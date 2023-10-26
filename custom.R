# layout_summary
setwd('C:/Users/Lenovo/Documents/projects/OfficeR')
if(!require(officer)) install.packages('officer')
library(officer)
library(ggplot2)

doc = read_pptx('template.pptx')
doc = on_slide(doc, index = 1)

ftitle = fp_text(
  font.family = 'SansSerif',
  bold = TRUE,
  color = '#222266',
  font.size = 40
)
ftitlepar = fp_par(text.align = 'center')

doc = ph_with(doc,
        value = fpar(ftext("Análise das Movimentações Econômicas de Biribi - ST", ftitle), fp_p = ftitlepar),
        location = ph_location(left = 0.15, top = 2, width = 13, height = 2)
      )

print(doc, 'custom.pptx')
