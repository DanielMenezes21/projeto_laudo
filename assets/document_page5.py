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
        nomes = item.get("nomes", [item.get("nome")]) if "nomes" in item else [item.get("nome")]
        cpfs = item.get("cpfs", [item.get("cpf")]) if "cpfs" in item else [item.get("cpf")]
        imovel = item.get("nome_imovel", "").strip()
        matricula = item.get("matricula", "").strip()
        chave = frozenset((n.strip(), c.strip()) for n, c in zip(nomes, cpfs))

        if chave:
            grupos[chave]["matriculas"].add(matricula)
            grupos[chave]["imoveis"].add(imovel)

    paragrafos = []

    for proprietarios, dados in grupos.items():
        nomes_cpfs = sorted(proprietarios)
        nomes = [n for n, _ in nomes_cpfs]
        cpfs = [c for _, c in nomes_cpfs]

        if len(nomes_cpfs) == 1:
            parte_proprietario = f"{nomes[0]}, inscrito sob o CPF nº {cpfs[0]}"
            verbo = "é o proprietário"
        else:
            parte_proprietario = ", ".join([f"{n}, inscrito sob o CPF nº {c}" for n, c in nomes_cpfs[:-1]])
            parte_proprietario += f" e {nomes_cpfs[-1][0]}, inscrito sob o CPF nº {nomes_cpfs[-1][1]}"
            verbo = "são os proprietários"

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
    heading = doc.add_heading("1 - SOLICITANTE", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")

    nomes_imoveis = sorted(set([m.get("nome_imovel", "").strip() for m in lista_matriculas if m.get("nome_imovel")]))
    cidade = lista_matriculas[0].get("municipio", "")
    estado = lista_matriculas[0].get("estado", "")

    if len(nomes_imoveis) == 1:
        texto = f"        Fomos solicitados pelo {solicitante}, para avaliar um imóvel rural, denominado {nomes_imoveis[0]}, localizado em {cidade} - {estado}."
    else:
        imoveis = ", ".join(nomes_imoveis[:-1]) + f" e {nomes_imoveis[-1]}"
        texto = f"        Fomos solicitados pelo {solicitante}, para avaliar os imóveis rurais, denominados {imoveis}, localizados em {cidade} - {estado}."

    doc.add_paragraph(texto)
    return doc

def texto_objetivo(doc, lista_matriculas):
    heading = doc.add_heading("2 - OBJETIVO", level=1)
    run = heading.runs[0]
    run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(" ")

    nomes_imoveis = sorted(set([m.get("nome_imovel", "").strip() for m in lista_matriculas if m.get("nome_imovel")]))

    if len(nomes_imoveis) == 1:
        texto = f"        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente ao imóvel {nomes_imoveis[0]}, localizado em #CIDADE_I - #ESTADO_I."
    else:
        imoveis = ", ".join(nomes_imoveis[:-1]) + f" e {nomes_imoveis[-1]}"
        texto = f"        O objetivo dessa peça técnica é aferir os valores de mercado e de liquidação forçada por meio do método comparativo de dados de mercado, referente aos imóveis {imoveis}, localizados em #CIDADE_I - #ESTADO_I."

    doc.add_paragraph(texto)
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