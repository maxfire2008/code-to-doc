import os
import docx
import time
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

# Measures the runtime of the program
start = time.time()

# Get the filepath of the folder
mainFolderPath = input("Enter the filepath of the folder: ")

# Get the filename
file_name = os.path.splitext(mainFolderPath)[0]

# Create a document and set the column into 2
document = Document()
section = document.sections[0]
section.top_margin = Inches(0.5)
section.bottom_margin = Inches(0.5)
section.left_margin = Inches(0.5)
section.right_margin = Inches(0.5)

# Create code style
styles = document.styles
code_style = styles.add_style("Code", WD_STYLE_TYPE.PARAGRAPH)
font = code_style.font
font.name = "Cascadia Code"
font.size = Pt(9)
# p_format = code_style.paragraph_format
# p_format.left_indent = Inches(0.5)
# p_format.first_line_indent = Inches(-0.5)

# Disable proofing (no spell check / grammar check)
rPr = code_style.element.get_or_add_rPr()
no_proof = OxmlElement("w:noProof")
rPr.append(no_proof)

# Open the files using their paths and transfer the content to docx
for path, subdirs, files in os.walk(mainFolderPath):
    for pathFile in files:
        pathFile = os.path.join(path, pathFile)
        fileName = os.path.basename(pathFile)
        srcFile = open(pathFile, "r+", encoding='UTF-8').read().splitlines()
        relName = os.path.relpath(pathFile, mainFolderPath)

        # Add a header [filename.php]
        head = document.add_heading(level=2)
        #head.paragraph_format.space_before = Pt(15)
        #head.paragraph_format.space_after = Pt(15)
        head = head.add_run(relName)
        #fhead = head.font
        #fhead.name = 'Arial'
        #fhead.size = Pt(10)
        #fhead.bold = True
        #fhead.italic = True
        #fhead.underline = True

        lcount_len = len(str(len(srcFile)))        

        # Add the codes
        for i, line in enumerate(srcFile):
            lnumber = str(i + 1).rjust(lcount_len, " ")
            line = lnumber + " " + line
            code = document.add_paragraph(style='Code')
            code.paragraph_format.space_after = 0
            code = code.add_run(line)

# Saving to docx file
file_name_docx = file_name + ".docx"
document.save(file_name_docx)

# Saving to pdf for faster viewing
file_name_pdf = file_name + ".pdf"
convert(file_name_docx, file_name_pdf)
print(f"{file_name_docx} and {file_name_pdf} were successfully created!")

end = time.time()
print(f"Program runtime: {end - start} seconds")
