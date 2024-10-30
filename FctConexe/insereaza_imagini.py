import os
import sys
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

# change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate(r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\Anunt UAT\Anunt UAT.docx")
placeholder_1 = InlineImage(
    doc, 'Placeholders/Placeholder_1.png', width=Cm(5), height=Cm(4))

context = {
    'placeholder_1': placeholder_1
}


doc.render(context)
doc.save(r"G:\Shared drives\Root\11. DATABASE\01. Automatizari avize\Executie\03.Pentru finalizare\Anunt UAT\Anunt UAT - cu.docx")
