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

def set_table_fixed_width(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

def title_imovel(doc):
    heading = doc.add_paragraph( style='Heading 1')
    run = heading.add_run("6 - IDENTIFICAÇÃO E CARACTERIZAÇÃO DO IMÓVEL AVALIANDO")
    run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph("  \n  ")

    return doc

def localizacao(doc, lista_dados_matriculas):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.1 - Localização")
    run.font.color.rgb = RGBColor(0, 0, 0)

    locais_set = set()
    coords_dict = {}

    for dados in lista_dados_matriculas:
        cidade = dados.get("cidade", "").strip()
        estado = dados.get("estado", "").strip()
        latitude = dados.get("latitude", "").strip()
        longitude = dados.get("longitude", "").strip()
        matricula = dados.get("matricula", "").strip()
        nome_imovel = dados.get("nome_imovel", "").strip()

        if cidade and estado:
            locais_set.add(f"{cidade} - {estado}")

        if latitude and longitude:
            chave = (latitude, longitude)
            if chave not in coords_dict:
                coords_dict[chave] = []

            if (matricula, nome_imovel) not in coords_dict[chave]:
                coords_dict[chave].append((matricula, nome_imovel))

    print("🔎 Coordenadas agrupadas por centróide:")
    for (lat, lon), lista in coords_dict.items():
        print(f"  Matrícula(s): {[m for m, _ in lista]} | Imóvel(is): {[n for _, n in lista]} | Latitude: {lat} | Longitude: {lon}")

    texto_locais = ", ".join(sorted(locais_set))
    run1 = doc.add_paragraph(f"Zona Rural, Município: #CIDADE_I - #ESTADO_I")
    run2 = doc.add_paragraph("")

    for (lat, lon), lista in coords_dict.items():
        matriculas = [m for m, _ in lista]
        imoveis = [n for _, n in lista]
        if len(matriculas) == 1:
            par = doc.add_paragraph(f"Matrícula nº {matriculas[0]}, imóvel {imoveis[0]}")
            par.runs[0].underline = True
            par2 = doc.add_paragraph("Coordenadas:")
            par2.runs[0].underline = True
            par3 = doc.add_paragraph(f"Latitude: {lat}")
            par4 = doc.add_paragraph(f"Longitude: {lon}")
        else:
            matriculas_str = ", ".join(matriculas[:-1]) + " e " + matriculas[-1]
            imoveis_str = ", ".join(imoveis[:-1]) + " e " + imoveis[-1]
            par = doc.add_paragraph(f"Matrículas nº {matriculas_str}, imóveis {imoveis_str} respectivamente")
            par.runs[0].underline = True
            par2 = doc.add_paragraph("Coordenadas (centróide compartilhado):")
            par2.runs[0].underline = True
            par3 = doc.add_paragraph(f"Latitude: {lat}")
            par4 = doc.add_paragraph(f"Longitude: {lon}")

    return doc

def acesso(doc, imagem_acesso=None):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.2 - ROTA DE ACESSO MATRÌCULA {mat}")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph("{rota_acesso}")
    for r in run1.runs:
        r.font.size = Pt(12)
    if imagem_acesso:
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_img.add_run("#IMAGEM_ACESSO")
            shape = run.add_picture(imagem_acesso, width=Cm(15), height=Cm(8.77))
            border_xml = (
                '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<pic:spPr>'
                '<a:ln w="9525">'  
                '<a:solidFill>'
                '<a:srgbClr val="000000"/>'  
                '</a:solidFill>'
                '<a:prstDash val="solid"/>'  
                '</a:ln>'
                '</pic:spPr>'
                '</pic:pic>'
            )

            pic = run._r.xpath('.//pic:pic')[0]
            pic.append(parse_xml(border_xml))
            
        except Exception as e:
            print(f"Erro ao adicionar imagem: {e}")
            par = doc.add_paragraph("[ESPAÇO PARA ACESSO]", style='Normal')
    else:
        par = doc.add_paragraph("[ESPAÇO PARA ACESSO]", style='Normal')

    return doc

def carac_reg(doc):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.3 - Caracterização da Região")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph("{caracterização}")
    return doc

def desc_imovel(doc):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4 - Descrição do imóvel")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph("Trata-se de um imóvel rural de Matrícula nº {n_matricula}, " \
    "com área total de {area_total} ha, destes, {p_reserva} são separados para Reserva Legal, " \
    "totalizando uma área de {area_reserva} ha, {observacao}, a sua Área de Preservação Permanente – APP " \
    "ocupa uma área {p_app}, totalizando {area_app} ha da integralidade do imóvel.")
    run2 = doc.add_paragraph("{descricao_atividade}")
    run3 = doc.add_paragraph("Uma melhor percepção do imóvel pode ser obtida através da tabela e das imagens a seguir:")
    run4 = doc.add_paragraph(" ")

    table = doc.add_table(rows=1, cols=4)
    table.allow_autofit = False
    table.style = 'Table Grid'
    table.width = Cm(16)

    table.rows[0].height = Cm(0.5)

    col_widths = [Cm(4), Cm(9), Cm(2.5), Cm(2.5)]

    for col_idx, width in enumerate(col_widths):
        for row in table.rows:
            row.cells[col_idx].width = width

    merged_cell = table.cell(0, 0).merge(table.cell(0, 3))

    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    merged_cell._tc.get_or_add_tcPr().append(shading)

    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' 
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    merged_cell._tc.get_or_add_tcPr().append(borders)

    p = merged_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("AVALIANDO")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)

    table1 = doc.add_table(rows=5, cols=4)
    table1.allow_autofit = False
    table1.width = Cm(16)
    set_table_fixed_width(table1)
    table1.style = 'Table Grid'

    table1.rows[0].height = Cm(0.5)
    table1.rows[1].height = Cm(0.5)
    table1.rows[2].height = Cm(0.5)
    table1.rows[3].height = Cm(0.5)
    table1.rows[4].height = Cm(0.5)

    col_widths = [Cm(4), Cm(8), Cm(2), Cm(2)]

    for col_idx, width in enumerate(col_widths):
        for row in table1.rows:
            row.cells[col_idx].width = width

    line1r1table1 = table1.rows[0].cells[0]
    line1r1table1.text = "Proprietário(s): "
    line1r1table1.paragraphs[0].runs[0].bold = True
    line1r1table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = line1r1table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line1r2table1 = table1.rows[0].cells[1]
    line1r2table1.text = "{nome}"
    line1r2table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tcPr = line1r2table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line1r3table1 = table1.rows[0].cells[2]
    line1r3table1.text = "Contato: "
    line1r3table1.paragraphs[0].runs[0].bold = True
    line1r3table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = line1r3table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line1r4table1 = table1.rows[0].cells[3]
    line1r4table1.text = "{contato}"
    line1r4table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = line1r4table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line2r1table1 = table1.rows[1].cells[0]
    line2r1table1.text = "Localização do imóvel: "
    line2r1table1.paragraphs[0].runs[0].bold = True
    line2r1table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = line2r1table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line2r2table1 = table1.rows[1].cells[1]
    line2r2table1.text = "{cit_est}"
    line2r2table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tcPr = line2r2table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line2r3table1 = table1.rows[1].cells[2]
    line2r3table1.text = "Situação: "
    line2r3table1.paragraphs[0].runs[0].bold = True
    line2r3table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = line2r3table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line2r4table1 = table1.rows[1].cells[3]
    line2r4table1.text = "{situacao}"
    line2r4table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = line2r4table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line3r1table1 = table1.rows[2].cells[0]
    line3r1table1.text = "Área do imóvel (ha): "
    line3r1table1.paragraphs[0].runs[0].bold = True
    line3r1table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = line3r1table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line3r2table1 = table1.rows[2].cells[1]
    line3r2table1.text = "{area_imovel}"
    line3r2table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tcPr = line3r2table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line3r3table1 = table1.rows[2].cells[2]
    line3r3table1.text = "Benfeitoria: "
    line3r3table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    line3r3table1.paragraphs[0].runs[0].bold = True
    tcPr = line3r3table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line3r4table1 = table1.rows[2].cells[3]
    line3r4table1.text = "{situacao}"
    line3r4table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = line3r4table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line4r1table1 = table1.rows[3].cells[0]
    line4r1table1.text = "Hidrografia: "
    line4r1table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    line4r1table1.paragraphs[0].runs[0].bold = True
    tcPr = line4r1table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line4r2table1 = table1.rows[3].cells[1]
    line4r2table1.text = "{hidrografia}"
    line4r2table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tcPr = line4r2table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line4r3table1 = table1.rows[3].cells[2]
    line4r3table1.text = "Data: "
    line4r3table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    line4r3table1.paragraphs[0].runs[0].bold = True
    tcPr = line4r3table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line4r4table1 = table1.rows[3].cells[3]
    line4r4table1.text = "{data_planilha}"
    line4r4table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = line4r4table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line5r1table1 = table1.rows[4].cells[0]
    line5r1table1.text = "Matrículas: "
    line5r1table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    line5r1table1.paragraphs[0].runs[0].bold = True
    tcPr = line5r1table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line5r2table1 = table1.rows[4].cells[1]
    line5r2table1.text = "{matricula}"
    line5r2table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    tcPr = line5r2table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line5r3table1 = table1.rows[4].cells[2]
    line5r3table1.text = "Outros: "
    line5r3table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    line5r3table1.paragraphs[0].runs[0].bold = True
    tcPr = line5r3table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    line5r4table1 = table1.rows[4].cells[3]
    line5r4table1.text = "{outros}"
    line5r4table1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = line5r4table1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    extra_table = doc.add_paragraph("[INSERIR_TABELA_AQUI]")

    run4

    table2 = doc.add_table(rows=3,cols=2)
    table2.autofit
    table2.width = Cm(16)
    set_table_fixed_width(table2)
    table2.style = 'Table Grid'

    table2.rows[0].height = Cm(0.5)
    table2.rows[2].height = Cm(0.5)

    titlecell = table2.rows[0].cells[0]
    titlecell.merge(table2.rows[0].cells[1])
    titlecell.text = "IMAGENS DE SATÉLITE DA PROPRIEDADE RURAL"
    titlecell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    titlecell.paragraphs[0].runs[0].bold = True
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    titlecell._tc.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' 
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    titlecell._tc.get_or_add_tcPr().append(borders)

    image_cell_1 = table2.rows[1].cells[0]
    p_img = image_cell_1.paragraphs[0]
    run_img = p_img.add_run()
    run_img.add_picture("captura_teste.png", width=Cm(8), height=Cm(4.84))
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

    description_cell_1 = table2.rows[2].cells[0]
    description_cell_1.text = "{descricao_1}"
    description_cell_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    image_cell_2 = table2.rows[1].cells[1]
    p_img = image_cell_2.paragraphs[0]
    run_img = p_img.add_run()
    run_img.add_picture("captura_teste.png", width=Cm(8), height=Cm(4.84))
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

    description_cell_2 = table2.rows[2].cells[1]
    description_cell_2.text = "{descricao_2}"
    description_cell_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    return doc

def declividade(doc, imagem_path=None):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4.1 - Declividade")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    if imagem_path:
        p = doc.add_paragraph()
        run = p.add_run("#IMAGEM_DECLIVIDADE")
        run.add_picture(imagem_path, width=Cm(14))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph = doc.add_paragraph("a área apresenta declividade #DECLIVIDADE_I")
    return doc

def hidrografia(doc, imagem_path=None):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4.2 - Hidrografia")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    if imagem_path:
        p = doc.add_paragraph()
        run = p.add_run("#IMAGEM_HIDROGRAFIA")
        run.add_picture(imagem_path, width=Cm(14))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph = doc.add_paragraph("a área apresenta hidrografia #HIDROGRAFIA_I")
    return doc

def pedologia(doc, imagem_path=None):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4.3 - Solo/Pedologia")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    if imagem_path:
        p = doc.add_paragraph()
        run = p.add_run("#IMAGEM_SOLOS")
        run.add_picture(imagem_path, width=Cm(14))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph = doc.add_paragraph("a pedologia da região é predominada por #TIPO_SOLO")
    paragraph2 = doc.add_paragraph("a área apresenta #DESCRICAO_SOLO")
    return doc

def uso_imovel(doc):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4.4 – Potencial de Utilização do Imóvel")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    paragraph = doc.add_paragraph("O imóvel é utilizado para #ATIVIDADE_IMOVEL")
    return doc

def benfeitoria(doc):
    heading = doc.add_paragraph( style='Heading 2')
    run = heading.add_run("6.4.5 – Benfeitorias")
    for run in heading.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

    paragraph = doc.add_paragraph("O imóvel possui as seguintes benfeitorias: ")
    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'

    cell_1_1 = table.cell(0, 0)
    cell_1_2 = table.cell(0, 1)
    cell_1_1.text = 'BENFEITORIA'
    cell_1_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_1_1.paragraphs[0].runs[0].bold = True
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell_1_1._tc.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>' 
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell_1_1._tc.get_or_add_tcPr().append(borders)

    cell_1_2.text = 'DESCRIÇÃO'
    cell_1_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_1_2.paragraphs[0].runs[0].bold = True
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="4EA65D"/>')
    cell_1_2._tc.get_or_add_tcPr().append(shading)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    cell_1_2._tc.get_or_add_tcPr().append(borders)

    cell_benfeitoria = table.cell(1, 0).merge(table.cell(3, 0))

    cell_dimensoes = table.cell(1, 1)
    par = cell_dimensoes.paragraphs[0]
    par.add_run('Comprimento: 20,0 m\n')
    par.add_run('Largura: 10,0 m\n')
    par.add_run('Altura: 10,0 m\n')
    par.add_run('Área Total: 200,0 m²')

    cell_desc = table.cell(2, 1)
    cell_desc.text = (
        'O galpão é ideal para armazenamento e organização de bags de fertilizantes e sementes para plantio, '
        'além de alocar os maquinários da propriedade.'
    )

    cell_ava = table.cell(3, 1)
    adicionar_estado_conservacao_com_tabela(cell_ava)

    return doc

def adicionar_estado_conservacao_com_tabela(cell):
    """Insere uma mini-tabela com estado de conservação dentro da célula"""
    p = cell.paragraphs[0]
    p.add_run('Estado de conservação:')

    nested_table = cell.add_table(rows=3, cols=7)
    nested_table.autofit = True
    nested_table.style = 'Table Grid'

    cell1_1 = nested_table.cell(0, 0)
    cell1_1.text = 'Bom'
    cell1_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_2 = nested_table.cell(0, 1)
    cell1_2.text = ' '
    cell1_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_2._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_3 = nested_table.cell(0, 2)
    cell1_3.text = ' '
    cell1_3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_3._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_4 = nested_table.cell(0, 3)
    cell1_4.text = ' '
    cell1_4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_5 = nested_table.cell(0, 4)
    cell1_5.text = 'Ruim'
    cell1_5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_6 = nested_table.cell(0, 5)
    cell1_6.text = ' '
    cell1_6.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_6._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell1_7 = nested_table.cell(0, 6)
    cell1_7.text = ' '
    cell1_7.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell1_7._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_1 = nested_table.cell(1, 0)
    cell2_1.text = ' '
    cell2_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_2 = nested_table.cell(1, 1)
    cell2_2.text = ' '
    cell2_2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_2._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
    cell2_3 = nested_table.cell(1, 2)
    cell2_3.text = 'Mediano'
    cell2_3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER 
    tcPr = cell2_3._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_4 = nested_table.cell(1, 3)
    cell2_4.text = ' '
    cell2_4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_4._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_5 = nested_table.cell(1, 4)
    cell2_5.text = ' '
    cell2_5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_5._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_6 = nested_table.cell(1, 5)
    cell2_6.text = ' '
    cell2_6.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_6._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell2_7 = nested_table.cell(1, 6)
    cell2_7.text = ' '
    cell2_7.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell2_7._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)

    cell3_1 = nested_table.cell(2, 0)
    cell3_1.merge(nested_table.cell(2, 6))
    cell3_1.text = ' '
    cell3_1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcPr = cell3_1._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders)
