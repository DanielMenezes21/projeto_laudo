from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import docx

def set_cell_background(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)

    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '8')
        border.set(qn('w:space'), '0')      
        border.set(qn('w:color'), '000000') 
        tblBorders.append(border)

    tblPr.append(tblBorders)

def set_row_height(row, height_cm):
    for cell in row.cells:
        cell.height = Pt(height_cm * 28.35)  
        cell.height_rule = WD_ALIGN_VERTICAL.CENTER  


def remove_cell_padding(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    cell_margin = OxmlElement('w:cellMar')
    cell_margin.set(qn('w:top'), '0')  
    cell_margin.set(qn('w:left'), '0')  
    cell_margin.set(qn('w:bottom'), '0')  
    cell_margin.set(qn('w:right'), '0')  
    tcPr.append(cell_margin)

def set_line_spacing(paragraph, new_line_spacing):
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(new_line_spacing)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
doc = Document()

table = doc.add_table(rows=3, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER  

set_table_borders(table)

cell_merged = table.cell(0, 0).merge(table.cell(0, 1))
cell_merged.text = 'Texto no fundo verde'
set_cell_background(cell_merged, '4EA65D')  
paragraph = cell_merged.paragraphs[0]
paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER

run = paragraph.runs[0]
font = run.font
font.name = 'Cambria'
font.size = Pt(11)
font.bold = True  

imagem1_path = 'captura_teste.png'  
imagem2_path = 'captura_teste.png'  

paragraph1 = table.cell(1, 0).paragraphs[0]
run1 = paragraph1.add_run()
run1.add_picture(imagem1_path, width=Cm(7), height=Cm(4.31))

paragraph1.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(1, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

paragraph2 = table.cell(1, 1).paragraphs[0]
run2 = paragraph2.add_run()
run2.add_picture(imagem2_path, width=Cm(7), height=Cm(4.31))

paragraph2.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(1, 1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

table.cell(2, 0).text = 'Texto simples 1'
table.cell(2, 1).text = 'Texto simples 2'

for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER  
            for run in paragraph.runs:
                font = run.font
                font.name = 'Cambria'
                font.size = Pt(11)

            set_line_spacing(paragraph, Pt(0)) 
            
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER  

set_row_height(table.rows[0], 0.5)  
set_row_height(table.rows[1], 4.31) 
set_row_height(table.rows[2], 0.5)

for row in table.rows:
    for cell in row.cells:
        remove_cell_padding(cell)

doc.save('tabela_formatada.docx')
