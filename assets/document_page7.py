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

def diag_mercado(doc):
    heading = doc.add_paragraph(style='Heading 1')
    run = heading.add_run("7 - Diagnóstico do Mercado")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    
    run1 = doc.add_paragraph("Considerando a retomada da atividade econômica que vem " \
    "ocorrendo em nosso país, com um incentivo dos governos municipal e federal na construção" \
    " de novas edificações residenciais, comerciais e industriais, o município encontra-se com um " \
    "desempenho normal, havendo na cidade um número significativo de transações imobiliárias, " \
    "com absorção considerada normal. A liquidez do imóvel avaliando é considerada como média, " \
    "estando o desempenho do mercado normal. ")

    return doc

def metodologia(doc):
    heading = doc.add_paragraph(style='Heading 1')
    run = heading.add_run("8 - Metodologia")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    run1 = doc.add_paragraph("O método adotado para a avaliação do imóvel é o " \
    "Direto de Dados de Mercado que identifica o valor do bem por meio de cálculos " \
    "estatísticos baseados em imóveis semelhantes ao avaliando dos valores de seus componentes.")
    run2 = doc.add_paragraph("Para a avaliação da área de terreno, procedemos a pesquisas " \
    "junto ao mercado imobiliário local e corretores atuantes que transacionam imóveis " \
    "semelhantes ao do objeto da presente avaliação.")
    run3 = doc.add_paragraph("Conforme pesquisa realizada na região do imóvel avaliando, " \
    "para apuração de valor venal de mercado de terrenos, verificamos ser possível a utilização " \
    "neste trabalho do Método Comparativo Direto de Dados de Mercado, que deve ter a preferência, " \
    "sempre que possível, de acordo com a recomendação constante da NBR 14.653-3, em seu item 8.1.1:")

    run_extra = doc.add_paragraph("")

    run4 = doc.add_paragraph("“... Para a identificação do valor de mercado, " \
    "sempre que possível preferir o Método Comparativo Direto de Dados de Mercado”.")
    run4.italic = True

    return doc

def metodo_comparativo(doc):
    heading = doc.add_paragraph(style='Heading2')
    run = heading.add_run("8.1 - Método Comparativo Direto de Dados de Mercado")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par = doc.add_paragraph()
    run1 = par.add_run("Conforme item 7.3.1 da NBR 14.653-1, a conceituação do método é a seguinte: ")
    run1_1 = par.add_run("“Identifica o custo do bem por meio de tratamento técnico dos atributos dos elementos comparáveis, constituintes da amostra”.")
    run1_1.italic = True
    run2 = doc.add_paragraph("É condição fundamental para aplicação deste método a existência de um conjunto de dados que possa ser tomada, estatisticamente, " \
    "como amostra do mercado imobiliário.")

    return doc

def aproveitamento(doc):
    heading = doc.add_paragraph(style='Heading2')
    run = heading.add_run("8.2 - Aproveitamento")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    
    par = doc.add_paragraph()
    run1 = par.add_run("O princípio que norteou o trabalho avaliatório " \
    "foi o do aproveitamento eficiente, determinado por análise do mercado imobiliário, " \
    "cujo conceito encontra-se assim definido pela ABNT NBR 14653-3: ")
    run1_1=par.add_run("“Aquele recomendável e tecnicamente possível para o local, numa data de referência, observada a tendência mercadológica nas circunvizinhanças, entre os diversos usos permitidos pela legislação pertinentes”.")
    return doc

def especificacao(doc):
    heading = doc.add_paragraph(style='Heading1')
    run = heading.add_run("9 - ESPECIFICAÇÃO DA AVALIAÇÃO")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    
    return doc

def grau_especificacao(doc):
    heading = doc.add_paragraph(style='Heading2')
    run = heading.add_run("9.1 - Grau de Especificação")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par = doc.add_paragraph()

    run1 = par.add_run("No desenvolvimento do presente trabalho foi aplicado " \
    "tratamento dos dados por homogeneização através de fatores, fundamentados por " \
    "estudos conforme o item 8.2.1.4.2, da norma em questão.")

    par2 = doc.add_paragraph()
    run2 = par2.add_run("TABELA 3 – Grau de fundamentação no caso de utilização do " \
    "tratamento por fatores – Item 9.2.2 – ABNT NBR 14653-3 ")
    run2.font.bold = True

    table = doc.add_table(rows=7, cols=5)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False

    table.rows[0].height = Cm(0.8)
    table.rows[1].height = Cm(0.8)
    table.rows[2].height = Cm(2.35)
    table.rows[3].height = Cm(3.70)
    table.rows[4].height = Cm(1.70)
    table.rows[5].height = Cm(2.20)
    table.rows[6].height = Cm(1.70)

    cell1_1 = table.cell(0, 0)
    cell1_1.merge(table.cell(1, 0))
    cell1_1.text = "ITEM"
    cell1_1.width = Cm(2.0)
    cell1_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell1_1._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell1_1._element.get_or_add_tcPr().append(borders)

    cell1_2 = table.cell(0, 1)
    cell1_2.merge(table.cell(1, 1))
    cell1_2.text = "DESCRIÇÃO"
    cell1_2.width = Cm(3.60)
    cell1_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell1_2._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell1_2._element.get_or_add_tcPr().append(borders)

    cell1_3 = table.cell(0, 2)
    cell1_3.merge(table.cell(0, 4))
    cell1_3.text = "GRAU"
    cell1_3.width = Cm(1.7)
    cell1_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell1_3._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell1_3._element.get_or_add_tcPr().append(borders)

    cell2_3 = table.cell(1, 2)
    cell2_3.text = "III"
    cell2_3.width = Cm(3.6)
    cell2_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_3._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_3._element.get_or_add_tcPr().append(borders)

    cell2_4 = table.cell(1, 3)
    cell2_4.text = "II"
    cell2_4.width = Cm(1.7)
    cell2_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_4._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_4._element.get_or_add_tcPr().append(borders)

    cell2_5 = table.cell(1, 4)
    cell2_5.text = "I"
    cell2_5.width = Cm(1.7)
    cell2_5.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_5._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_5._element.get_or_add_tcPr().append(borders)

    cell3_1 = table.cell(2, 0)
    cell3_1.text = "1"
    cell3_1.width = Cm(2.0)
    cell3_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_1._element.get_or_add_tcPr().append(borders)

    cell3_2 = table.cell(2, 1)
    cell3_2.text = "Quantidade mínima de dados efetivamente utilizados"
    cell3_2.width = Cm(3.60)
    cell3_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_2._element.get_or_add_tcPr().append(borders)

    cell3_3 = table.cell(2, 2)
    cell3_3.text = "12"
    cell3_3.width = Cm(2.8)
    cell3_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_3._element.get_or_add_tcPr().append(borders)

    cell3_4 = table.cell(2, 3)
    cell3_4.text = "5"
    cell3_4.width = Cm(2.8)
    cell3_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_4._element.get_or_add_tcPr().append(borders)

    cell3_5 = table.cell(2, 4)
    cell3_5.text = "3"
    cell3_5.width = Cm(2.8)
    cell3_5.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_5._element.get_or_add_tcPr().append(borders)

    cell4_1 = table.cell(3, 0)
    cell4_1.text = "2"
    cell4_1.width = Cm(2.0)
    cell4_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell4_1._element.get_or_add_tcPr().append(borders)

    cell4_2 = table.cell(3, 1)
    cell4_2.text = "Apresentação dos dados"
    cell4_2.width = Cm(3.60)
    cell4_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell4_2._element.get_or_add_tcPr().append(borders)

    cell4_3 = table.cell(3, 2)
    cell4_3.text = "Atributos relativos a todos os dados e variáveis analisados na modelagem, com foto"
    cell4_3.width = Cm(2.8)
    cell4_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell4_3._element.get_or_add_tcPr().append(borders)

    cell4_4 = table.cell(3, 3)
    cell4_4.text = "Atributos relativos a todos os dados e variáveis analisados na modelagem"
    cell4_4.width = Cm(2.8)
    cell4_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell4_4._element.get_or_add_tcPr().append(borders)

    cell4_5 = table.cell(3, 4)
    cell4_5.text = "Atributos relativos aos dados e variáveis efetivamente utilizados no modelo"
    cell4_5.width = Cm(2.8)
    cell4_5.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell4_5._element.get_or_add_tcPr().append(borders)

    cell5_1 = table.cell(4, 0)
    cell5_1.text = "3"
    cell5_1.width = Cm(2.0)
    cell5_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell5_1._element.get_or_add_tcPr().append(borders)

    cell5_2 = table.cell(4, 1)
    cell5_2.text = "Origem dos fatores de homogeneização"
    cell5_2.width = Cm(3.60)
    cell5_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell5_2._element.get_or_add_tcPr().append(borders)

    cell5_3 = table.cell(4, 2)
    cell5_3.text = "Estudos embasados em metodologia científica"
    cell5_3.width = Cm(2.8)
    cell5_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell5_3._element.get_or_add_tcPr().append(borders)

    cell5_4 = table.cell(4, 3)
    cell5_4.text = "Publicações"
    cell5_4.width = Cm(2.8)
    cell5_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell5_4._element.get_or_add_tcPr().append(borders)

    cell5_5 = table.cell(4, 4)
    cell5_5.text = "Análise do avaliador"
    cell5_5.width = Cm(2.8)
    cell5_5.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell5_5._element.get_or_add_tcPr().append(borders)

    cell6_1 = table.cell(5, 0)
    cell6_1.text = "4"
    cell6_1.width = Cm(2.0)
    cell6_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell6_1._element.get_or_add_tcPr().append(borders)

    cell6_2 = table.cell(5, 1)
    cell6_2.text = "Intervalo admissível de ajuste para o conjunto de fatores"
    cell6_2.width = Cm(3.60)
    cell6_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell6_2._element.get_or_add_tcPr().append(borders)

    cell6_3 = table.cell(5, 2)
    cell6_3.text = "0,80 a 1,25"
    cell6_3.width = Cm(2.8)
    cell6_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell6_3._element.get_or_add_tcPr().append(borders)

    cell6_4 = table.cell(5, 3)
    cell6_4.text = "0,70 a 1,40"
    cell6_4.width = Cm(2.8)
    cell6_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell6_4._element.get_or_add_tcPr().append(borders)

    cell6_5 = table.cell(5, 4)
    cell6_5.text = "0,50 a 2,00"
    cell6_5.width = Cm(2.8)
    cell6_5.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell6_5._element.get_or_add_tcPr().append(borders)

    cell7_1 = table.cell(6, 0)
    cell7_1.merge(table.cell(6, 4))
    cell7_1.text = "ª No caso de utilização de menos de cinco dados de mercado, " \
    "o intervalo admissível de ajuste é de 0,80 a 1,25,pois é desejável que, com um número " \
    "menor de dados de mercado, a amostra seja heterogênea."
    cell7_1.width = Cm(10.0)
    cell7_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell7_1._element.get_or_add_tcPr().append(borders)

    run3 = doc.add_paragraph("O atendimento a cada exigência do Grau I vale um ponto, " \
    "do Grau II, dois pontos, do Grau III, três pontos. O enquadramento global do Laudo deve " \
    "considerar a soma de pontos obtidos para o conjunto de itens, atendendo a tabela 3.")

    run4 = doc.add_paragraph("Neste trabalho foram contabilizados 10 pontos, " \
    "correspondentes à soma de pontos dos itens atingidos e acima destacados.")

    parrun5 = doc.add_paragraph()
    run5 = parrun5.add_run("TABELA 4 – Enquadramento do laudo segundo seu grau de fundamentação " \
    "no caso de utilização de tratamento por fatores – Item 9.2.2.2 – ABNT NBR 14653-3")
    run5.font.bold = True

    table2 = doc.add_table(rows=3, cols=4)
    table2.style = 'Table Grid'
    table2.autofit = False
    table2.allow_autofit = False

    table2.rows[0].height = Cm(0.8)
    table2.rows[1].height = Cm(0.8)
    table2.rows[2].height = Cm(1.5)

    table2cell1_1 = table2.cell(0, 0)
    table2cell1_1.text = "GRAUS"
    table2cell1_1.width = Cm(3.9)
    table2cell1_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    table2cell1_1._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell1_1._element.get_or_add_tcPr().append(borders)

    table2cell1_2 = table2.cell(0, 1)
    table2cell1_2.text = "III"
    table2cell1_2.width = Cm(3.9)
    table2cell1_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    table2cell1_2._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell1_2._element.get_or_add_tcPr().append(borders)

    table2cell1_3 = table2.cell(0, 2)
    table2cell1_3.text = "II"
    table2cell1_3.width = Cm(3.9)
    table2cell1_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    table2cell1_3._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell1_3._element.get_or_add_tcPr().append(borders)
    table2cell1_4 = table2.cell(0, 3)
    table2cell1_4.text = "I"
    table2cell1_4.width = Cm(3.9)
    table2cell1_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    table2cell1_4._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell1_4._element.get_or_add_tcPr().append(borders)

    table2cell2_1 = table2.cell(1, 0)
    table2cell2_1.text = "Pontos mínimos"
    table2cell2_1.width = Cm(3.9)
    table2cell2_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell2_1._element.get_or_add_tcPr().append(borders)

    table2cell2_2 = table2.cell(1, 1)
    table2cell2_2.text = "13"
    table2cell2_2.width = Cm(3.9)
    table2cell2_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell2_2._element.get_or_add_tcPr().append(borders)

    table2cell2_3 = table2.cell(1, 2)
    table2cell2_3.text = "8"
    table2cell2_3.width = Cm(3.9)
    table2cell2_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell2_3._element.get_or_add_tcPr().append(borders)

    table2cell2_4 = table2.cell(1, 3)
    table2cell2_4.text = "5"
    table2cell2_4.width = Cm(3.9)
    table2cell2_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell2_4._element.get_or_add_tcPr().append(borders)

    table2cell3_1 = table2.cell(2, 0)
    table2cell3_1.text = "Itens Obrigatórios"
    table2cell3_1.width = Cm(3.9)
    table2cell3_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell3_1._element.get_or_add_tcPr().append(borders)

    table2cell3_2 = table2.cell(2, 1)
    table2cell3_2.text = "2, 4 e 5 no grau III e os demais no mínimo no grau II"
    table2cell3_2.width = Cm(3.9)
    table2cell3_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell3_2._element.get_or_add_tcPr().append(borders)

    table2cell3_3 = table2.cell(2, 2)
    table2cell3_3.text = "2, 4 e 5 no grau II e os demais no mínimo no grau I"
    table2cell3_3.width = Cm(3.9)
    table2cell3_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell3_3._element.get_or_add_tcPr().append(borders)

    table2cell3_4 = table2.cell(2, 3)
    table2cell3_4.text = "Todos no mínimo no grau I"
    table2cell3_4.width = Cm(3.9)
    table2cell3_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    table2cell3_4._element.get_or_add_tcPr().append(borders)

    parrun6 = doc.add_paragraph()
    run6 = parrun6.add_run("GRAU DE FUNDAMENTAÇÃO ")
    run6.font.bold = True
    run6.underline = True

    run_extra = doc.add_paragraph(" ")


    run7 = doc.add_paragraph("Com base nos parâmetros especificados pelas tabelas 3 e 4 da ABNT, NBR 14653-3, embora tenhamos alcançado 06 pontos e atendido as exigências da norma no grau ll todos os itens atendem à exigência da norma no grau II, consequentemente, o trabalho avaliatório será enquadrado no ")
    run7_1 = run7.add_run("GRAU II")
    run7_1.font.bold = True

    return doc

def grau_precisao(doc):
    """
    Create a document with a table and text related to the 'Grau de Precisão' section.
    """
    heading = doc.add_paragraph(style='Heading2')
    run = heading.add_run("9.2 - Grau de Precisão")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    par = doc.add_paragraph()
    run1 = par.add_run("TABELA 5 – Grau de precisão nos casos de utilização de modelos" \
    " de regressão linear ou do tratamento por fatores – Item 9.2.3 ¬– ABNT NBR 14653-3 ")
    run1.font.bold = True

    table = doc.add_table(rows=3, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False

    table.rows[0].height = Cm(0.8)
    table.rows[1].height = Cm(0.8)
    table.rows[2].height = Cm(0.8)

    cell1_1 = table.cell(0, 0)
    cell1_1.merge(table.cell(1, 0))
    cell1_1.text = "DESCRIÇÃO"
    cell1_1.width = Cm(8.2)
    cell1_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell1_1._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell1_1._element.get_or_add_tcPr().append(borders)

    cell1_2 = table.cell(0, 1)
    cell1_2.merge(table.cell(0, 3))
    cell1_2.text = "GRAU"
    cell1_2.width = Cm(7.5)
    cell1_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell1_2._element.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell1_2._element.get_or_add_tcPr().append(borders)

    cell2_1 = table.cell(1, 1)
    cell2_1.text = "I"
    cell2_1.width = Cm(2.5)
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_1._element.get_or_add_tcPr().append(shading)
    cell2_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_1._element.get_or_add_tcPr().append(borders)

    cell2_2 = table.cell(1, 2)
    cell2_2.text = "II"
    cell2_2.width = Cm(2.5)
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_2._element.get_or_add_tcPr().append(shading)
    cell2_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_2._element.get_or_add_tcPr().append(borders)

    cell2_3 = table.cell(1, 3)
    cell2_3.text = "III"
    cell2_3.width = Cm(2.5)
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell2_3._element.get_or_add_tcPr().append(shading)
    cell2_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell2_3._element.get_or_add_tcPr().append(borders)

    cell3_1 = table.cell(2, 0)
    cell3_1.text = "Amplitude do intervalo de confiança de 80% em torno " \
    "da estimativa de tendência central"
    cell3_1.width = Cm(8.2)
    cell3_1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_1._element.get_or_add_tcPr().append(borders)

    cell3_2 = table.cell(2, 1)
    cell3_2.text = "≤ 30%"
    cell3_2.width = Cm(2.5)
    cell3_2.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_2._element.get_or_add_tcPr().append(borders)

    cell3_3 = table.cell(2, 2)
    cell3_3.text = "≤ 40%"
    cell3_3.width = Cm(2.5)
    cell3_3.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_3._element.get_or_add_tcPr().append(borders)

    cell3_4 = table.cell(2, 3)
    cell3_4.text = "≤ 50%"
    cell3_4.width = Cm(2.5)
    cell3_4.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell3_4._element.get_or_add_tcPr().append(borders)

    return doc

def grau_precisao2(doc):
    """
    Create a document with a table and text related to the 'Grau de Precisão' section.
    """
    heading = doc.add_paragraph(style='Heading2')
    run = heading.add_run("9.3 - GRAU DE PRECISÃO")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    run1 = doc.add_paragraph("Considerando os parâmetros especificados na " \
    "tabela 1 da ABNT NBR 14653-3, referente ao grau de precisão, o presente trabalho " \
    "está enquadrado no ")
    run2 = run1.add_run("GRAU II")
    run2.font.bold = True

    return doc

def resultado(doc):
    """
    Create a document with a table and text related to the 'Resultado' section.
    """
    heading = doc.add_paragraph(style='Heading1')
    run = heading.add_run("10 - RESULTADO DA AVALIAÇÃO")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    parrun1 = doc.add_paragraph()
    run1 = parrun1.add_run("Conforme a NBR 14653-1, Avaliação de Bens, Parte-1: " \
    "Procedimentos Gerais, item 3.1.9 e NBR 14653-2, Avaliação de Bens, Parte-3: Imóveis Rurais, " \
    "item A5 (Anexo A), o Campo de Arbítrio é o intervalo compreendido entre o valor máximo e o " \
    "mínimo dos preços homogeneizados, efetivamente utilizados no tratamento, limitado a 15% do valor " \
    "calculado, dentro do qual se pode arbitrar, pelo avaliador, o valor mais representativo do bem.")

    parrun2 = doc.add_paragraph()
    run2 = parrun2.add_run("De acordo com a análise do diagnóstico de mercado apurado e " \
    "levantamentos realizados, com base na identificação da realidade mercadológica da " \
    "região onde se encontra situado o imóvel, os valores mínimo, médio e máximo que espelham o " \
    "valor venal de mercado foram discriminados de acordo com a amplitude do intervalo de confiança " \
    "de 80% em torno do valor central da estimativa. ")

    parrun3 = doc.add_paraagraph()
    run3 = parrun3.add_run("[INSERIR_VALORES_AQUI]")

    return doc

