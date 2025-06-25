from docx import Document
from docx.shared import Pt, Cm
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Twips
from collections import defaultdict

def adicionar_espaco(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)  
    p.paragraph_format.space_after = Pt(0)   
    run = p.add_run(" ")
    run.font.name = 'Calibri'
    run.font.size = Pt(12) 

    return doc

def gerar_texto_proprietarios(lista_matriculas):
    # Agrupa por proprietário
    grupo_por_proprietario = defaultdict(lambda: {
        "cpf": "",
        "imoveis": set(),
        "matriculas": set()
    })

    for item in lista_matriculas:
        nome = item.get("nome", "").strip()
        cpf = item.get("cpf", "").strip()
        imovel = item.get("nome_imovel", "").strip()
        matricula = item.get("matricula", "").strip()

        if nome and cpf:
            grupo_por_proprietario[nome]["cpf"] = cpf
            grupo_por_proprietario[nome]["imoveis"].add(imovel)
            grupo_por_proprietario[nome]["matriculas"].add(matricula)

    parágrafos = []

    for nome, dados in grupo_por_proprietario.items():
        cpf = dados["cpf"]
        imoveis = sorted(dados["imoveis"])
        matriculas = sorted(dados["matriculas"])

        # Determina se é um ou mais imóveis
        texto_imovel = (
            f"o imóvel rural denominado {imoveis[0]}"
            if len(imoveis) == 1 else
            f"os imóveis rurais denominados {', '.join(imoveis[:-1])} e {imoveis[-1]}"
        )

        # Determina se é uma ou mais matrículas
        texto_matricula = (
            f"a matrícula de nº {matriculas[0]}"
            if len(matriculas) == 1 else
            f"as matrículas de nº {', '.join(matriculas[:-1])} e {matriculas[-1]}"
        )

        parágrafo = (
            f"        Em conformidade com o exposto em {texto_matricula}, "
            f"{nome}, inscrito sob o CPF nº {cpf}, é o proprietário de {texto_imovel}."
        )
        parágrafos.append(parágrafo)

    return "\n\n".join(parágrafos)


def texto_solicitante(doc):
    heading = doc.add_heading("1 - SOLICITANTE", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Fomos solicitados pelo #SOLICITANTE, para avaliar um imóvel rural, denominado #NOME_IMOVEL, localizado em #MUNICIPIO - #ESTADO ")

    return doc

def texto_objetivo(doc):
    heading = doc.add_heading("2 - OBJETIVO", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente ao imóvel #NOME_IMOVEL, localizado em #MUNICIPIO - #ESTADO")

    return doc

def texto_finalidade(doc):
    heading = doc.add_heading("3 - FINALIDADE", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Garantia bancária")

    return doc

def texto_proprietario(doc, lista_matriculas):
    heading = doc.add_heading("4 - PROPRIETÁRIO", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    texto = gerar_texto_proprietarios(lista_matriculas)
    doc.add_paragraph(texto)
    return doc

def texto_ressalvas(doc):
    heading = doc.add_heading("5 - PRESSUPOSTOS, RESSALVAS E FATORES IMPORTANTES", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")
    doc.add_paragraph("        Este Laudo fundamenta-se no que estabelecem as normas técnicas da ABNT"\
    "Avaliação de Bens, NBR 14653 – Parte 1 (Procedimentos Gerais/Revisão 2019) e Parte 3"\
    "(Imóveis Rurais/Revisão 2011), e baseia-se na documentação fornecida referente ao imóvel localizado"\
    "em #MUNICIPIO - #ESTADO, situação na qual o #TRATAMENTO #PROPONENTE solicita a avaliação do mesmo." \
    " Quanto às edificações e benfeitorias existentes no imóvel são considerados os quantitativos" \
    "de projetos existentes (se existirem), informações constatadas in loco quando da vistoria ao imóvel, " \
    "realizada em {data_av} e sendo, dessa forma, adotadas na presente avaliação como oficiais, por premissa," \
    " consideradas como válidas.\n "
    "    Também, utilizamos como referência no decorrer dos trabalhos elementos documentais " \
    "e informações prestadas por terceiros, admitidas como confiáveis, corretas e de boa fé.")

    return doc