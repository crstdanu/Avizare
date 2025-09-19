import os
import sys
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

# change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate(r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\Cerere Transelectrica.docx")
placeholder_1 = InlineImage(
    doc, 'Placeholders/Placeholder_1.png', width=Mm(50), height=Mm(35))



context = {
    'Placeholder_1': placeholder_1,
}


doc.render(context)
doc.save(r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\MODELE\IS\27. Aviz TransElectrica\Cerere Transelectrica.docx")
