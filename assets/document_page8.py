from docx import Document
from docx.shared import Pt, Cm
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips
import os
import win32com.client
from win32com.client.gencache import EnsureDispatch
from win32com.client import Dispatch
import num2words
from docx.shared import Inches

def encerramento(doc):
    heading = doc.add_paragraph(style='Heading 1')
    run = heading.add_run("11 - ENCERRAMENTO:")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par = doc.add_paragraph()
    par.add_run("Ante o exposto e de acordo com a análise técnica realizada, informamos que o valor Venal mais representativo para o imóvel em questão é de ")
    run1 = par.add_run("{valor_medio} ({valor_extenso})")
    run1.bold = True
    par.add_run(". Já o valor de liquidação forçada obtido foi de ")
    run2 = par.add_run("{valor_liq}")
    run2.bold = True

    doc.add_paragraph("\n\n")

    linha1 = doc.add_paragraph()
    linha1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    linha1.add_run("_________________________________________")
    assinatura1 = doc.add_paragraph()
    assinatura1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    assinatura1.add_run("Eng. Marcos Felipe Oliveira Sousa\nCREA 333267/D-TO")

    doc.add_paragraph("\n")

    linha2 = doc.add_paragraph()
    linha2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    linha2.add_run("_________________________________________")
    assinatura2 = doc.add_paragraph()
    assinatura2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    assinatura2.add_run("Eng. Luhan Marcos Pereira Lustosa\nCREA 326186/D-TO")

    doc.add_paragraph("\n")

    linha3 = doc.add_paragraph()
    linha3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    linha3.add_run("_________________________________________")
    assinatura3 = doc.add_paragraph()
    assinatura3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    assinatura3.add_run("D’AGRO SOLUÇÕES EM AGRONEGÓCIOS\nCNPJ 60.087.243/0001-30\nCREA 1000102812")

    return doc

def inserir_marcadagua_so_na_secao(path_docx, path_img, secao=2):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(path_docx)
    if doc.Sections.Count >= secao:
        section = doc.Sections(secao)
        rng = section.Range
        shape = doc.Shapes.AddPicture(
            FileName=os.path.abspath(path_img),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=section.PageSetup.PageWidth,
            Height=section.PageSetup.PageHeight,
            Anchor=rng
        )
        shape.WrapFormat.Type = 3  # Behind text
        shape.LockAspectRatio = False
        shape.RelativeHorizontalPosition = 0
        shape.RelativeVerticalPosition = 0
        shape.Left = 0
        shape.Top = 0
        doc.Save()
    doc.Close()
    word.Quit()

def inserir_caixa_texto(doc):
    table1 = doc.add_table(rows=1, cols=1)
    cell1 = table1.cell(0, 0)
    cell1.text = "ANEXOS"
    cell1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table1.alignment = WD_ALIGN_PARAGRAPH.CENTER 

    tc1 = cell1._tc
    tcPr1 = tc1.get_or_add_tcPr()

    for run in cell1.paragraphs[0].runs:
        run.font.size = Pt(22) 
        run.font.color.rgb = RGBColor(255, 255, 255)  

    doc.add_paragraph()

    table2 = doc.add_table(rows=3, cols=1)
    anexos = [
        "ANEXO I – RELATÓRIO FOTOGRÁFICO",
        "ANEXO II – DOCUMENTAÇÃO DO IMÓVEL",
        "ANEXO III –  PARÂMETROS DE AVALIAÇÃO E MEMORIAL DE CÁLCULO"
    ]
    for i, texto in enumerate(anexos):
        cell = table2.cell(i, 0)
        cell.text = texto
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(12)  
            run.font.color.rgb = RGBColor(255, 255, 255)  

    for i in range(3):
        tc = table2.cell(i, 0)._tc
        tcPr = tc.get_or_add_tcPr()
        """shd = parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls('w')))
        tcPr.append(shd)"""

    table2.alignment = WD_ALIGN_PARAGRAPH.LEFT 

    return doc

def anexos_fotos(doc):
    heading = doc.add_paragraph(style='Heading 1')
    run = heading.add_run("ANEXO I - RELATÓRIO FOTOGRÁFICO")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    return doc

def anexo_doc(doc):
    heading = doc.add_paragraph(style='Heading 1')
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run = heading.add_run("ANEXO II - DOCUMENTAÇÂO DO IMÓVEL")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run1 = par.add_run("CERTIDÃO DE INTEIRO TEOR")
    run1.bold = True

    par2 = doc.add_paragraph()
    par2.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run2 = par2.add_run("#SUBSTITUIR_CIT")

    par3 = doc.add_paragraph()
    par3.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run3 = par3.add_run("RECIBO DO CADASTRO AMBIENTAL RURAL – CAR")
    run3.bold = True

    par4 = doc.add_paragraph()
    par4.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run4 = par4.add_run("#SUBSTITUIR_CAR")

    return doc

def anexo_parametros(doc):
    heading = doc.add_paragraph(style='Heading 1')
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run = heading.add_run("ANEXO III – PARÂMETROS DE AVALIAÇÃO E MEMORIAL DE CÁLCULO")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par1 = doc.add_paragraph()
    run1 = par1.add_run("Benfeitoria")
    run1.bold = True

    par2 = doc.add_paragraph()
    run2 = par2.add_run("[INSERIR_BENFEITORIA_AQUI]")

    par3 = doc.add_paragraph()
    run3 = par3.add_run("Depreciação de benfeitorias")
    run3.bold = True

    par4 = doc.add_paragraph()
    run4 = par4.add_run("[INSERIR_DEPRECIACAO_AQUI]")

    par5 = doc.add_paragraph()
    run5 = par5.add_run("•	Adequada")
    run5.bold = True
    run5_1 = par5.add_run("= edificação está perfeitamente adequada à sua utilização; " \
    "está 100% aproveitada e/ou funcional e/ou utilizada, considerando o imóvel e a região " \
    "num período de um ano agrícola;")