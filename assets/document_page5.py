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
    grupos = defaultdict(lambda: {"matriculas": set(), "imoveis": set()})

    for item in lista_matriculas:
        proprietario = item.get("proprietario", "").strip()
        cpf = item.get("cpf", "").strip()
        imovel = item.get("nome_imovel", "").strip()
        matricula = item.get("matricula", "").strip()
        chave = (proprietario, cpf)
        if proprietario:
            grupos[chave]["matriculas"].add(matricula)
            grupos[chave]["imoveis"].add(imovel)

    paragrafos = []
    for (nome, cpf), dados in grupos.items():
        if nome:
            parte_proprietario = f"{nome}, inscrito sob o CPF nº {cpf}" if cpf else nome
            verbo = "é o proprietário"
            matriculas = sorted(dados["matriculas"])
            texto_matricula = (
                f"na matrícula de nº {matriculas[0]}" if len(matriculas) == 1
                else f"nas matrículas de nº {', '.join(matriculas[:-1])} e {matriculas[-1]}"
            )
            imoveis = sorted(dados["imoveis"])
            texto_imovel = (
                f"do imóvel rural denominado {imoveis[0]}" if len(imoveis) == 1
                else f"dos imóveis rurais denominados {', '.join(imoveis[:-1])} e {imoveis[-1]}"
            )
            paragrafo = (
                f"        Em conformidade com o exposto {texto_matricula}, "
                f"{parte_proprietario}, {verbo} {texto_imovel}."
            )
            paragrafos.append(paragrafo)

    return "\n\n".join(paragrafos)

def texto_solicitante(doc, solicitante, lista_matriculas):
    heading = doc.add_paragraph(style='Heading1')
    run = heading.add_run("1 - SOLICITANTE")
    run.font.color.rgb = RGBColor(0, 0, 0)
    par = doc.add_paragraph()
    runpar = par.add_run("")

    nomes_imoveis = sorted(set([m.get("nome_imovel", "").strip() for m in lista_matriculas if m.get("nome_imovel")]))

    if len(nomes_imoveis) == 1:
        texto = f"        Fomos solicitados pelo {solicitante}, para avaliar um imóvel rural, denominado {nomes_imoveis[0]}, localizado em #CIDADE_I - #ESTADO_I."
    else:
        imoveis = ", ".join(nomes_imoveis[:-1]) + f" e {nomes_imoveis[-1]}"
        texto = f"        Fomos solicitados pelo {solicitante}, para avaliar os imóveis rurais, denominados {imoveis}, localizados em #CIDADE_I - #ESTADO_I."

    par = doc.add_paragraph()
    runpar2 = par.add_run(texto)
    return doc

def texto_objetivo(doc, lista_matriculas):
    heading = doc.add_paragraph( style='Heading1')
    run = heading.add_run("2 - OBJETIVO")
    run.font.color.rgb = RGBColor(0, 0, 0)
    par = doc.add_paragraph()
    runpar = par.add_run("")

    nomes_imoveis = sorted(set([m.get("nome_imovel", "").strip() for m in lista_matriculas if m.get("nome_imovel")]))

    if len(nomes_imoveis) == 1:
        texto = f"        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente ao imóvel {nomes_imoveis[0]}, localizado em #CIDADE_I - #ESTADO_I."
    else:
        imoveis = ", ".join(nomes_imoveis[:-1]) + f" e {nomes_imoveis[-1]}"
        texto = f"        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente aos imóveis {imoveis}, localizados em #CIDADE_I - #ESTADO_I."

    par = doc.add_paragraph()
    runpar2 = par.add_run(texto)
    return doc


def texto_finalidade(doc):
    heading = doc.add_paragraph(style='Heading1')
    run = heading.add_run("3 - FINALIDADE")
    run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph(" ")
    run2 = doc.add_paragraph("        Garantia bancária")

    return doc

def texto_proprietario(doc, lista_matriculas):
    heading = doc.add_paragraph( style='Heading1')
    run = heading.add_run("4 - PROPRIETÁRIO")
    run.font.color.rgb = RGBColor(0, 0, 0)
    run1 =doc.add_paragraph(" ")
    texto = gerar_texto_proprietarios(lista_matriculas)
    run2 = doc.add_paragraph(texto)
    return doc

def texto_ressalvas(doc):
    heading = doc.add_paragraph( style='Heading1')
    run = heading.add_run("5 - PRESSUPOSTOS, RESSALVAS E FATORES IMPORTANTES")
    run.font.color.rgb = RGBColor(0, 0, 0)
    run1 = doc.add_paragraph(" ")
    run2 = doc.add_paragraph("        Este Laudo fundamenta-se no que estabelecem as normas técnicas da ABNT"\
    "Avaliação de Bens, NBR 14653 – Parte 1 (Procedimentos Gerais/Revisão 2019) e Parte 3"\
    "(Imóveis Rurais/Revisão 2011), e baseia-se na documentação fornecida referente ao imóvel localizado"\
    "em #CIDADE_I - #ESTADO_I, situação na qual o #TRATAMENTO #PROPONENTE solicita a avaliação do mesmo." \
    " Quanto às edificações e benfeitorias existentes no imóvel são considerados os quantitativos" \
    "de projetos existentes (se existirem), informações constatadas in loco quando da vistoria ao imóvel, " \
    "realizada em {data_av} e sendo, dessa forma, adotadas na presente avaliação como oficiais, por premissa," \
    " consideradas como válidas.\n "
    "    Também, utilizamos como referência no decorrer dos trabalhos elementos documentais " \
    "e informações prestadas por terceiros, admitidas como confiáveis, corretas e de boa fé.")

    return doc