from docx import Document
from docx.shared import Pt, Cm
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Twips

def adicionar_espaco(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)  
    p.paragraph_format.space_after = Pt(0)   
    run = p.add_run(" ")
    run.font.name = 'Calibri'
    run.font.size = Pt(12) 

    return doc

def texto_solicitante(doc):
    heading = doc.add_heading("1 - SOLICITANTE", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Fomos solicitados pelo {prop} {nome}, para avaliar um imóvel rural, denominado {imovel}, localizado em {cid_est} ")

    return doc

def texto_objetivo(doc):
    heading = doc.add_heading("2 - OBJETIVO", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente ao imóvel {imovel}, localizado em {cid_est}")

    return doc

def texto_finalidade(doc):
    heading = doc.add_heading("3 - FINALIDADE", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Garantia bancária")

    return doc

def texto_proprietario(doc):
    heading = doc.add_heading("4 - PROPRIETÁRIO", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Em conformidade com o exposto na matrícula de nº 154.725, o {trat} {nome} inscrito sob o CPF nº {cpf}, é o proprietário do imóvel rural denominado {imovel}, {sit_civil}")

    return doc

def texto_ressalvas(doc):
    heading = doc.add_heading("5 - PRESSUPOSTOS, RESSALVAS E FATORES IMPORTANTES", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Este Laudo fundamenta-se no que estabelecem as normas técnicas da ABNT"\
    "Avaliação de Bens, NBR 14653 – Parte 1 (Procedimentos Gerais/Revisão 2019) e Parte 3"\
    "(Imóveis Rurais/Revisão 2011), e baseia-se na documentação fornecida referente ao imóvel localizado"\
    "em {cit_est}, situação na qual o {trat} {sol} solicita a avaliação do mesmo." \
    " Quanto às edificações e benfeitorias existentes no imóvel são considerados os quantitativos" \
    "de projetos existentes (se existirem), informações constatadas in loco quando da vistoria ao imóvel, " \
    "realizada em {data_av} e sendo, dessa forma, adotadas na presente avaliação como oficiais, por premissa," \
    " consideradas como válidas.\n "
    "    Também, utilizamos como referência no decorrer dos trabalhos elementos documentais " \
    "e informações prestadas por terceiros, admitidas como confiáveis, corretas e de boa fé.")

    return doc