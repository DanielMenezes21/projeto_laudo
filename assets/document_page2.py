from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml
import os
from docx.oxml.ns import nsdecls
from docx.shared import Twips

def colorir_celula(cell, cor_hex="009933"):
        tcPr = cell._tc.get_or_add_tcPr()
        for el in tcPr.findall('.//w:shd', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            tcPr.remove(el)
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{cor_hex}"/>')
        tcPr.append(shading)

def cm_to_twips(cm):
    """Converte centímetros para Twips (1cm = 567 Twips)."""
    return Twips(cm * 567)

LARGURA = cm_to_twips(18)

def adicionar_linha_fina(doc):
    """
    Adiciona um caractere do tamanho 1
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)  
    p.paragraph_format.space_after = Pt(0)   
    run = p.add_run("w")
    run.font.name = 'Calibri'
    run.font.size = Pt(1) 

    return doc

def configurar_documento():
    """Configura as propriedades básicas do documento"""
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    style = doc.styles['Normal']
    style.paragraph_format.space_before = Cm(0)
    style.paragraph_format.space_after = Cm(0)
    style.paragraph_format.line_spacing = 1
    
    return doc

def criar_titulo(doc, dados):
    """Cria o título principal com fundo verde com dados específicos"""
    table = doc.add_table(rows=1, cols=1)
    usable_width = LARGURA
    table.allow_autofit = True
    table.width = usable_width
    table.style = 'Table Grid'

    cell = table.cell(0, 0)

    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell._tc.get_or_add_tcPr().append(shading)

    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' 
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell._tc.get_or_add_tcPr().append(borders)

    matricula = dados.get("matricula", "")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"FICHA CADASTRAL {matricula}")
    run.bold = True
    run.font.size = Pt(12)

    return doc

def criar_secao_valor(doc, dados):
    """Cria a seção de Valor Total do Imóvel com formatação específica"""
    usable_width = LARGURA
    table = doc.add_table(rows=3, cols=5)
    table.allow_autofit = True  
    table.width = usable_width  
    table.style = 'Table Grid'

    table.rows[0].height = Cm(0.76) 
    table.rows[1].height = Cm(0.19)  
    table.rows[2].height = Cm(0.19)  

    title_cell = table.rows[0].cells[0]
    title_cell.merge(table.rows[0].cells[4])  
    
    title_cell.text = "1. VALOR TOTAL DO IMÓVEL"
    tcPr = title_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>' 
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    title_cell.paragraphs[0].runs[0].bold = True
    title_cell.paragraphs[0].runs[0].font.size = Pt(12)

    valor_total = dados.get("valor_total","")
    matricula = dados.get("matricula","")
    row = table.rows[1].cells
    row[0].text = "VALOR TOTAL"
    row[0].paragraphs[0].runs[0].bold = True

    row[1].text = f"{valor_total}"
    row[1].paragraphs[0].runs[0].bold = True
    
    row[2].text = "MATRÍCULA"
    row[2].paragraphs[0].runs[0].bold = True

    row[3].text = f"{matricula}"
    row[3].paragraphs[0].runs[0].bold = True

    row[4].text = ""
    row[4].paragraphs[0].runs[0].bold = True

    for cell_idx in [0, 4]:
        tcPr = table.rows[1].cells[cell_idx]._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            '<w:top w:val="nil"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="nil"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tcBorders>'
        )
        tcPr.append(borders)

    for cell_idx in [0, 2]:
        tcPr = table.rows[1].cells[cell_idx]._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            '<w:top w:val="nil"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="nil"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tcBorders>'
        )
        tcPr.append(borders)

    valor_liq_total = dados.get("valor_liq_total","")
    row = table.rows[2].cells
    row[0].text = "LIQUIDAÇÃO"
    row[0].paragraphs[0].runs[0].bold = True

    row[1].text = f"{valor_liq_total}"
    row[1].paragraphs[0].runs[0].bold = True

    row[2].merge(row[4])

    for cell_idx in [2,3]:
        tcPr = table.rows[2].cells[cell_idx]._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            '<w:top w:val="nil"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tcBorders>'
        )
        tcPr.append(borders)

    tcPr = table.rows[2].cells[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    
    return doc

def criar_secao_identificacao(doc):
    """Cria a seção de Identificação do Imóvel com 4 linhas e 3 colunas"""
    usable_width = LARGURA
    table = doc.add_table(rows=5, cols=3)
    table.allow_autofit = True 
    table.width = usable_width
    table.style = 'Table Grid'

    table.rows[0].height = Cm(0.7) 
    table.rows[1].height = Cm(0.5)  
    table.rows[2].height = Cm(0.2)  
    table.rows[3].height = Cm(0.5)  
    table.rows[4].height = Cm(0.2)  

    title_cell = table.rows[0].cells[0]
    title_cell.merge(table.rows[0].cells[2])
    title_cell.text = "2. IDENTIFICAÇÃO DO IMÓVEL"
    title_cell.paragraphs[0].runs[0].bold = True
    title_cell.paragraphs[0].runs[0].font.size = Pt(12)
    title_cell.width = LARGURA

    tcPr = title_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>' 
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row2 = table.rows[1].cells
    row2[0].text = "IDENTIFICAÇÃO"
    row2[0].paragraphs[0].runs[0].bold = True
    row2[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row2[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    
    row2[1].text = "#NOME_IMOVEL"
    row2[1].paragraphs[0].runs[0].bold = True
    row2[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    row2[2].text = ''
    row2[2].paragraphs[0].runs[0].bold = True
    row2[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row2[2]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row3 = table.rows[2].cells
    row3[0].merge(row3[2])  

    tcPr = row3[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4 = table.rows[3].cells
    row4[0].text = "MUNICÍPIO"
    row4[0].paragraphs[0].runs[0].bold = True
    row4[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row4[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    
    row4[1].text = "#CIDADE_I - #ESTADO_I"
    row4[1].paragraphs[0].runs[0].bold = True
    row4[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    row4[2].text = ""
    row4[2].paragraphs[0].runs[0].bold = True
    row4[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row4[2]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row5 = table.rows[4].cells[0]
    row5.merge(table.rows[4].cells[2])
    paragraph = row5.paragraphs[0]
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run("|")
    run.font.size = Pt(1)
    tcPr = row5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' +
        '<w:top w:val="nil"/>' +
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    for row in [1, 2, 3]:  
        table.rows[row].cells[2].text = ""
    
    return doc

def criar_secao_croqui(doc, imagem_path=None):
    """Cria a seção do Croqui de Localização com borda na imagem"""
    table = doc.add_table(rows=1, cols=1)
    usable_width = LARGURA
    table.allow_autofit = True
    table.width = usable_width
    table.style = 'Table Grid'

    table.rows[0].height = Cm(0.5)  

    title_cell = table.rows[0].cells[0]
    title_cell.text = "3. CROQUI DE LOCALIZAÇÃO DO IMÓVEL"
    title_cell.paragraphs[0].runs[0].bold = True
    title_cell.paragraphs[0].runs[0].font.size = Pt(12)

    tcPr = title_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    if imagem_path and os.path.exists(imagem_path):
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_img.add_run()
            run.add_picture(imagem_path, width=Cm(15), height=Cm(10.61))
        except Exception as e:
            print(f"⚠ Erro ao adicionar imagem do croqui: {e}")
            par = doc.add_paragraph(style='Normal')
            runpar = par.add_run("[ESPAÇO PARA CROQUI]")
    else:
        par = doc.add_paragraph( style='Normal')
        runpar = par.add_run("[ESPAÇO PARA CROQUI]")

    return doc

def geometria_terreno(doc):
    """Cria a seção de Geometria do Terreno com polígonos"""
    table = doc.add_table(rows=3, cols=5)
    usable_width = LARGURA
    table.allow_autofit = True
    table.width = usable_width
    table.style = 'Table Grid'

    table.rows[0].height = Cm(1.01)  
    table.rows[1].height = Cm(0.56)
    table.rows[2].height = Cm(0.2)  

    title_cell = table.rows[0].cells[0]
    title_cell.merge(table.rows[0].cells[4])
    title_cell.text = "GEOMETRIA DO TERRENO"
    title_cell.paragraphs[0].runs[0].bold = True
    title_cell.paragraphs[0].runs[0].font.size = Pt(12)

    tcPr = title_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row = table.rows[1].cells
    row[0].text = "POLÍGONO REGULAR"
    row[0].paragraphs[0].runs[0].font.size = Pt(12)
    row[0].width = Cm(4.46)

    tcPr = row[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)    

    row[1].text = " "
    row[1].paragraphs[0].runs[0].font.size = Pt(12)
    row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row[1].width = Cm(0.46)
    row[1].height = Cm(0.56)

    tcPr = row[1]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    row[2].text = "POLÍGONO IRREGULAR"
    row[2].paragraphs[0].runs[0].font.size = Pt(12)
    row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    row[2].width = Cm(8.80)

    tcPr = row[2]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row[3].text = " "
    row[3].paragraphs[0].runs[0].font.size = Pt(12)
    row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row[3].width = Cm(0.46)
    row[3].height = Cm(0.56)

    tcPr = row[3]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row[4].text = " "
    row[4].paragraphs[0].runs[0].font.size = Pt(12)
    row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row[4].width = Cm(2.5)

    tcPr = row[4]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row1 = table.rows[2].cells[0]
    row1.merge(table.rows[2].cells[4])  
    paragraph = row1.paragraphs[0]
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    run = paragraph.add_run("|")
    run.font.size = Pt(1)

    tcPr = row1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    return doc, table

def criar_secao_caracteristicas(doc, dados):
    """Cria a seção de Características do Terreno"""
    table = doc.add_table(rows=9, cols=14)
    usable_width = LARGURA
    table.allow_autofit = True
    table.width = usable_width
    table.style = 'Table Grid'

    table.rows[0].height = Cm(0.42)
    table.rows[1].height = Cm(1.01)
    table.rows[2].height = Cm(1)
    table.rows[3].height = Cm(0.6)
    table.rows[4].height = Cm(0.05)
    table.rows[5].height = Cm(0.42)
    table.rows[6].height = Cm(0.42)
    table.rows[7].height = Cm(0.42)
    table.rows[8].height = Cm(0.05)

    checkboxes_af = dados.get("checkboxes_af", {})
    checkboxes_pares = dados.get("checkboxes_pares", {})
    checkboxes_superficie = dados.get("checkboxes_superficie", {})

    cor_hex = "4EA65D"

    title_cell = table.rows[0].cells[0]
    title_cell.merge(table.rows[0].cells[13])
    title_cell.text = "4. CARACTERÍSTICAS DO TERRENO"
    title_cell.paragraphs[0].runs[0].bold = True
    title_cell.paragraphs[0].runs[0].font.size = Pt(12)

    tcPr = title_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row1 = table.rows[1].cells[0]
    row1.merge(table.rows[1].cells[1])
    row1.text = "ÁREA TOTAL"
    row1.paragraphs[0].runs[0].font.size = Pt(12)
    row1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    row1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    row1.width = Cm(2)

    tcPr = row1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row1cell_2 = table.rows[1].cells[2]
    row1cell_2.text = "106,2710"
    row1cell_2.paragraphs[0].runs[0].font.size = Pt(12)
    row1cell_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1cell_2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    row1cell_2.width = Cm(2.7)

    tcPr = row1cell_2._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row1cell_3 = table.rows[1].cells[3]
    row1cell_3.merge(table.rows[1].cells[8])
    row1cell_3.text = "ÁREA CONSTRUÍDA (m²)"
    row1cell_3.paragraphs[0].runs[0].font.size = Pt(12)
    row1cell_3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    row1cell_3.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    row1cell_3.width = Cm(4.6)

    tcPr = row1cell_3._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )

    tcPr.append(borders)

    row1cell_4 = table.rows[1].cells[9]
    row1cell_4.merge(table.rows[1].cells[10])
    row1cell_4.text = "0,00"
    row1cell_4.paragraphs[0].runs[0].font.size = Pt(12)
    row1cell_4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1cell_4.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    row1cell_4.width = Cm(2.7)

    tcPr = row1cell_4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row1cell_5 = table.rows[1].cells[11]
    row1cell_5.merge(table.rows[1].cells[13])
    row1cell_5.text = ""
    row1cell_5.paragraphs[0].runs[0].font.size = Pt(12)
    row1cell_5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1cell_5.width = Cm(1)

    tcPr = row1cell_5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row2 = table.rows[2].cells[0]
    row2.merge(table.rows[2].cells[13])
    row2.text = "DECLIVIDADE"
    row2.paragraphs[0].runs[0].font.size = Pt(12)
    row2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    row2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    tcPr = row2._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3 = table.rows[3].cells
    row3[0].text="A"
    row3[0].paragraphs[0].runs[0].font.size = Pt(12)
    row3[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3[0].width = Cm(2.6)

    tcPr = row3[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3[1].text = " "
    row3[1].paragraphs[0].runs[0].font.size = Pt(12)
    row3[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3[1].width = Cm(0.4)

    tcPr = row3[1]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3[2].text = "B"
    row3[2].paragraphs[0].runs[0].font.size = Pt(12)
    row3[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3[2].width = Cm(2.6)

    tcPr = row3[2]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3[3].text = " "
    row3[3].paragraphs[0].runs[0].font.size = Pt(12)
    row3[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3[3].width = Cm(0.4)

    tcPr = row3[3]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_4 = table.rows[3].cells[4]
    row3cell_4.merge(table.rows[3].cells[5])
    row3cell_4.text = "C"
    row3cell_4.paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row3cell_4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_5 = table.rows[3].cells
    row3cell_5[6].text = " "
    row3cell_5[6].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_5[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_5[6].width = Cm(0.4)

    tcPr = row3cell_5[6]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_6 = table.rows[3].cells
    row3cell_6[7].text = "D"
    row3cell_6[7].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_6[7].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_6[7].width = Cm(2.2)

    tcPr = row3cell_6[7]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_7 = table.rows[3].cells
    row3cell_7[8].text = " "
    row3cell_7[8].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_7[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_7[8].width = Cm(0.4)

    tcPr = row3cell_7[8]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    
    row3cell_8 = table.rows[3].cells
    row3cell_8[9].text = "E"
    row3cell_8[9].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_8[9].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_8[9].width = Cm(2.6)

    tcPr = row3cell_8[9]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_9 = table.rows[3].cells
    row3cell_9[10].text = " "
    row3cell_9[10].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_9[10].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_9[10].width = Cm(0.4)

    tcPr = row3cell_9[10]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_10 = table.rows[3].cells
    row3cell_10[11].text = "F"
    row3cell_10[11].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_10[11].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_10[11].width = Cm(2.6)

    tcPr = row3cell_10[11]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_11 = table.rows[3].cells
    row3cell_11[12].text = " "
    row3cell_11[12].paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_11[12].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_11[12].width = Cm(0.4)

    tcPr = row3cell_11[12]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row3cell_12 = table.rows[3].cells[13]
    row3cell_12.text = " "
    row3cell_12.paragraphs[0].runs[0].font.size = Pt(12)
    row3cell_12.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row3cell_12.width = Cm(1.2)

    tcPr = row3cell_12._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row_extra = table.rows[4].cells[0]
    row_extra.merge(table.rows[4].cells[13])
    paragraph = row_extra.paragraphs[0]
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    run = paragraph.add_run("|")
    run.font.size = Pt(1)
    tcPr = row_extra._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' +
        '<w:top w:val="nil"/>' +
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
        '<w:bottom w:val="nil"/>' +
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4 = table.rows[5].cells
    row4[0].text="AB"
    row4[0].paragraphs[0].runs[0].font.size = Pt(12)
    row4[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4[0].width = Cm(2.6)

    tcPr = row4[0]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4[1].text = " "
    row4[1].paragraphs[0].runs[0].font.size = Pt(12)
    row4[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4[1].width = Cm(0.4)

    tcPr = row4[1]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4[2].text = "BA"
    row4[2].paragraphs[0].runs[0].font.size = Pt(12)
    row4[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4[2].width = Cm(2.6)

    tcPr = row4[2]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4[3].text = " "
    row4[3].paragraphs[0].runs[0].font.size = Pt(12)
    row4[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4[3].width = Cm(0.4)

    tcPr = row4[3]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_4 = table.rows[5].cells[4]
    row4cell_4.merge(table.rows[5].cells[5])
    row4cell_4.text = "BC"
    row4cell_4.paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row4cell_4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_5 = table.rows[5].cells
    row4cell_5[6].text = " "
    row4cell_5[6].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_5[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_5[6].width = Cm(0.4)

    tcPr = row4cell_5[6]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_6 = table.rows[5].cells
    row4cell_6[7].text = "CB"
    row4cell_6[7].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_6[7].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_6[7].width = Cm(2.2)

    tcPr = row4cell_6[7]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_7 = table.rows[5].cells
    row4cell_7[8].text = " "
    row4cell_7[8].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_7[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_7[8].width = Cm(0.4)

    tcPr = row4cell_7[8]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    
    row4cell_8 = table.rows[5].cells
    row4cell_8[9].text = "CD"
    row4cell_8[9].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_8[9].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_8[9].width = Cm(2.6)

    tcPr = row4cell_8[9]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_9 = table.rows[5].cells
    row4cell_9[10].text = " "
    row4cell_9[10].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_9[10].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_9[10].width = Cm(0.4)

    tcPr = row4cell_9[10]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_10 = table.rows[5].cells
    row4cell_10[11].text = "DC"
    row4cell_10[11].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_10[11].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_10[11].width = Cm(2.6)

    tcPr = row4cell_10[11]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_11 = table.rows[5].cells
    row4cell_11[12].text = " "
    row4cell_11[12].paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_11[12].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_11[12].width = Cm(0.4)

    tcPr = row4cell_11[12]._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row4cell_12 = table.rows[5].cells[13]
    row4cell_12.text = " "
    row4cell_12.paragraphs[0].runs[0].font.size = Pt(12)
    row4cell_12.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row4cell_12.width = Cm(1.2)

    tcPr = row4cell_12._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row5 = table.rows[6].cells[0]
    row5.merge(table.rows[6].cells[13])
    row5.text="SUPERFÍCIE DO SOLO"
    row5.paragraphs[0].runs[0].font.size = Pt(12)
    row5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    tcPr = row5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_1 = table.rows[7].cells[0]
    row6cell_1.merge(table.rows[7].cells[2])
    row6cell_1.text = "SECO"
    row6cell_1.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row6cell_1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_2 = table.rows[7].cells[3]
    row6cell_2.text = " "
    row6cell_2.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row6cell_2.width = Cm(0.4)

    tcPr = row6cell_2._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_3 = table.rows[7].cells[4]
    row6cell_3.merge(table.rows[7].cells[7])
    row6cell_3.text = "ÚMIDO"
    row6cell_3.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = row6cell_3._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell4 = table.rows[7].cells[8]
    row6cell4.text = " "
    row6cell4.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row6cell4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_5 = table.rows[7].cells[9]
    row6cell_5.merge(table.rows[7].cells[11])
    row6cell_5.text = "ALAGADIÇO"
    row6cell_5.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row6cell_5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_6 = table.rows[7].cells[12]
    row6cell_6.text = " "
    row6cell_6.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_6.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row6cell_6.width = Cm(0.4)

    tcPr = row6cell_6._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row6cell_7 = table.rows[7].cells[13]
    row6cell_7.text = " "
    row6cell_7.paragraphs[0].runs[0].font.size = Pt(12)
    row6cell_7.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcPr = row6cell_7._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    row7 = table.rows[8].cells[0]
    row7.merge(table.rows[8].cells[13])
    row7.width = LARGURA
    paragraph = row7.paragraphs[0]
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    run = paragraph.add_run("|")
    run.font.size = Pt(1)
    tcPr = row7._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' +
        '<w:top w:val="nil"/>' +
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' +
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    af_map = {"A": (3, 1), "B": (3, 3), "C": (3, 6), "D": (3, 8), "E": (3, 10), "F": (3, 12)}
    for letra, (row, col) in af_map.items():
        if checkboxes_af.get(letra):
            colorir_celula(table.rows[row].cells[col], cor_hex)

    pares_map = {"AB": (5, 1), "BA": (5, 3), "BC": (5, 6), "CB": (5, 8), "CD": (5, 10), "DC": (5, 12)}
    for par, (row, col) in pares_map.items():
        if checkboxes_pares.get(par):
            colorir_celula(table.rows[row].cells[col], cor_hex)

    superficie_map = {"seco": (7, 3), "umido": (7, 8), "alagadiço": (7, 12)}
    for nome, (row, col) in superficie_map.items():
        if checkboxes_superficie.get(nome):
            colorir_celula(table.rows[row].cells[col], cor_hex)

    return doc