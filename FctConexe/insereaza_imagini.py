import os
import sys
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

# change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate(
    r"E:\NEW_python\RGT\StudiiFezabilitate\Avize\modele_cereri\03. bacau\Cerere RAJA.docx")


placeholder_1 = InlineImage(
    doc, 'Placeholders/Placeholder_1.png', width=Mm(30), height=Mm(30))
placeholder_2 = InlineImage(
    doc, 'Placeholders/Placeholder_2.png', width=Mm(30), height=Mm(15))

context = {
    'placeholder_1': placeholder_1,
    'placeholder_2': placeholder_2,
}


doc.render(context)
doc.save(r"E:\NEW_python\RGT\StudiiFezabilitate\Avize\modele_cereri\03. bacau\01. Aviz RAJA\Cerere RAJA.docx")
