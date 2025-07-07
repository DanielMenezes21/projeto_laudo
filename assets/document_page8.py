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
from docx.enum.section import WD_ORIENT
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

    par6 = doc.add_paragraph()
    run6 = par6.add_run("•	Inadequada")
    run6.bolod = True
    run6_1 = par6.add_run("= edificação está parcialmente adequada à sua utilização; " \
    "aproximadamente 75% de sua capacidade é aproveitada e/ou funcional e/ou utilizada, " \
    "considerando o imóvel e a região num período de um ano agrícola;")

    par7 = doc.add_paragraph()
    run7 = par7.add_run("•	Superada")
    run7.bold = True
    run7_1 = par7.add_run("= edificação está superada, considerando as recomendações " \
    "técnicas atuais, mas aproximadamente 50% de sua capacidade ainda é aproveitada e/ou " \
    "funcional e/ou utilizada, considerando o imóvel e a região num período de um ano agrícola;")

    par8 = doc.add_paragraph()
    run8 = par8.add_run("•	Total")
    run8.bold = True
    run8_1 = par8.add_run("edificação não tem utilidade nenhuma, servindo apenas como " \
    "fonte de material usado; 20% aproveitada e/ou funcional e/ou utilizada, considerando o " \
    "imóvel e a região num período de um ano agrícola.")

    par9 = doc.add_paragraph()
    run9 = par9.add_run("Uso do solo")
    run9.bold = True

    par10 = doc.add_paragraph()
    run10 = par10.add_run("[INSERIR_CLASSE_AQUI]")

    par11 = doc.add_paragraph()
    run11 = par11.add_run("Classificação das terras quanto a aptidão de acordo com O Manual " \
    "brasileiro para levantamento da capacidade de uso da terra – Escritório Técnico " \
    "de Agricultura- Brasil /Estados Unidos 1971:")

    par12 = doc.add_paragraph()
    run12 = par12.add_run("•	CLASSE I ")
    run12.bold = True
    run12_1 = par12.add_run("- São terras que têm nenhuma ou somente muito pequenas limitações " \
    "permanentes ou riscos de depauperamento. São próprias para culturas anuais climaticamente " \
    "adaptadas, com produção de colheitas entre médias e elevadas, sem práticas ou medidas especiais " \
    "de conservação do solo. Normalmente, são solos profundos, de fácil mecanização, com boa retenção " \
    "de umidade no perfil e fertilidade de média a alta. São áreas planas ou com declividades muito " \
    "suaves, sem riscos de inundação e sem grandes restrições climáticas. Não há afloramentos de rocha, " \
    "nem o lençol de água é permanentemente elevado ou qualquer outra condição que possa prejudicar o uso " \
    "de máquinas agrícolas. Dependendo de bons sistemas de manejo, podem mesmo ser cultivadas com plantas " \
    "que facilitem a erosão, como o algodão, milho ou mandioca, plantadas em linhas retas, sem perigo " \
    "apreciável de erosão acelerada. ")

    par13 = doc.add_paragraph()
    run13 = par13.add_run("•	CLASSE II ")
    run13.bold = True
    run13_1 = par13.add_run("- Consiste em terras que têm limitações moderadas para o seu uso. Estão sujeitas " \
    "a riscos moderados de depauperamento, mas são terras boas, que podem ser cultivadas desde que lhes sejam " \
    "aplicadas práticas especiais de conservação do solo, de fácil execução, para produção segura e permanente " \
    "de colheitas entre médias e elevadas, de culturas anuais adaptadas à região. A declividade já pode ser " \
    "suficiente para provocar enxurradas e erosão. Em terras planas, podem requerer drenagem, porém sem " \
    "necessidades de práticas complexas de manutenção dos drenos. Podem enquadrar-se nessa classe também terras " \
    "que não tenham excelente capacidade de retenção de água. Cada uma dessas limitações requer cuidados " \
    "especiais, como aração e plantio em contorno, plantas de cobertura, cultura em faixas, controle de água, " \
    "proteção contra enxurradas advindas de glebas vizinhas, além das práticas comuns já referidas para a " \
    "classe l, como rotações de cultura e aplicações de corretivos e fertilizantes. ")

    par14 = doc.add_paragraph()
    run14 = par14.add_run("•	CLASSE III ")
    run14.bold = True
    run14_1 = par14.add_run("- São terras que, quando cultivadas sem cuidados especiais, são sujeitas a " \
    "severos riscos de depauperamento, principalmente no caso de culturas anuais. Requerem medidas intensas " \
    "e complexas de conservação do solo, a fim de poderem ser cultivadas segura e permanentemente, com produção " \
    "média a elevada, de culturas anuais adaptadas. Os principais fatores limitantes são a declividade "
    "(moderada), drenagem eficiente, escassez de água no solo (regiões semiáridas não irrigadas) e " \
    "pedregosidade. Frequentemente, essas limitações restringem muito a escolha das espécies a serem cultivadas, " \
    "ou à época do plantio ou operações de preparo e cultivo do solo. ")

    par15 = doc.add_paragraph()
    run15 = par15.add_run("•	CLASSE IV ")
    run15.bold = True
    run15_1 = par15.add_run("- São terras que têm riscos ou limitações permanentes muito severas " \
    "quando usadas para culturas anuais. Os solos podem ter fertilidade natural boa ou razoável, mas não são " \
    "adequados para cultivos intensivos e contínuos. Usualmente, devem ser mantidos com pastagens, " \
    "mas podem ser suficientemente boas para certos cultivos ocasionais (na proporção de um ano de cultivo " \
    "para cada quatro a seis de pastagens) ou para algumas culturas anuais, porém com cuidados muito especiais. " \
    "Tais terras podem ser caracterizadas pelos seguintes aspectos: declive íngreme, erosão severa, obstáculos " \
    "físicos, como pedregosidade ou drenagem muito deficiente, baixa produtividade, ou outras condições que " \
    "as tornem impróprias para o cultivo motomecanizado regular. Em algumas regiões onde a escassez de chuvas " \
    "seja muito sentida, de tal maneira a não serem seguras as culturas sem irrigação, as terras deverão ser " \
    "classificadas na Classe IV.")

    par16 = doc.add_paragraph()
    run16 = par16.add_run("•	CLASSE V ")
    run16.bold = True
    run16_1 = par16.add_run("- São terras planas ou com declives suaves praticamente livres de erosão, " \
    "mas impróprias para serem exploradas com culturas anuais, podendo, com segurança, ser apropriadas " \
    "para pastagens, florestas, ou mesmo para algumas culturas permanentes, sem a aplicação de técnicas " \
    "especiais. Embora apresentando praticamente planas e não sujeitas à erosão, não são adaptadas para " \
    "exploração com culturas anuais comuns, em razão de impedimentos permanentes, tais como muito baixa " \
    "capacidade de armazenamento de água, encharcamento (sem possibilidade de ser corrigido), adversidade " \
    "climática, frequente risco de inundação, pedregosidade ou afloramento de rochas. Em alguns casos " \
    "é possível o cultivo exclusivo de arroz; mesmo assim com risco de insucesso pelas limitações advindas, " \
    "principalmente, do risco de inundação. O solo, entretanto, tem poucas limitações de qualquer espécie, " \
    "para uso de pastagens ou silvicultura. Podem necessitar de alguns tratos para produções satisfatórias, " \
    "tanto de forragens como de arbustos e árvores. Entretanto, se tais tratos forem dispensados, não serão " \
    "sujeitas à erosão acelerada. Por isso, podem ser usadas permanentemente sem práticas especiais de controle " \
    "de erosão ou de proteção do solo. ")

    par17 = doc.add_paragraph()
    run17 = par17.add_run("•	CLASSE VI ")
    run17.bold = True
    run17_1 = par17.add_run("- Terras impróprias para culturas anuais, mas que podem ser usadas para " \
    "produção de certos cultivos permanentes úteis, como pastagens, florestas e algumas culturas permanentes " \
    "protetoras do solo, como seringueira e cacau, desde que adequadamente manejadas. O uso com pastagens ou " \
    "culturas permanentes protetoras devem ser feito com restrições moderadas, com práticas especiais de " \
    "conservação do solo, uma vez que, mesmo sob esse tipo de vegetação, são medianamente susceptíveis de " \
    "danificação pelos fatores de depauperamento do solo. Normalmente as limitações que apresentam, são em " \
    "razão da declividade excessiva ou pequena profundidade do solo, ou presença de pedras impedindo emprego " \
    "de máquinas agrícolas. Quando a pluviosidade da região é adequada para culturas, as limitações das " \
    "classes VI residem, em geral, na declividade excessiva, na pequena profundidade do solo ou na " \
    "pedregosidade. Nas regiões semiáridas, a escassez de umidade, muitas vezes, é a principal razão " \
    "para o enquadramento da terra na classe VI. ")

    par18 = doc.add_paragraph()
    run18 = par18.add_run("•	CLASSE VII ")
    run18.bold = True
    run18_1 = par18.add_run("- Terra que, por serem sujeitas a muitas limitações permanentes, além de " \
    "serem impróprias para culturas anuais, apresentam severas limitações, mesmo para certas culturas " \
    "permanentes protetoras do solo, pastagens e florestas. Sendo altamente susceptíveis de danificação, " \
    "exigem severas restrições de uso, com práticas especiais. Normalmente, são muito íngremes, " \
    "erodidas, pedregosas ou com solos muito rasos, ou ainda com deficiência de água muito grande. " \
    "Os cuidados necessários a elas são semelhantes aos aplicáveis à classe VI, com a diferença de poder " \
    "ser necessário maior número de práticas conservacionistas, ou que estas tenham que ser mais intensivas a " \
    "fim de prevenir ou diminuir os danos por erosão. Requerem cuidados extremos para controle da erosão. " \
    "Seu uso, tanto para pastoreio como para produção de madeira, requer sempre cuidados especiais.")

    par19 = doc.add_paragraph()
    run19 = par19.add_run("•	CLASSE VIII ")
    run19.bold = True
    run19_1 = par19.add_run("- Terras impróprias para serem utilizadas em qualquer tipo de cultivo, " \
    "inclusive de florestas comerciais ou para produção de qualquer outra forma de vegetação permanente " \
    "de valor econômico. Prestam-se apenas para proteção e abrigo da fauna e flora silvestre, para fins " \
    "de recreação e turismo ou de armazenamento de água em açudes. Consistem, em geral, em áreas extremamente " \
    "áridas, ou acidentadas, ou pedregosas, ou encharcadas (sem possibilidade de pastoreio ou drenagem " \
    "artificial), ou severamente erodidas ou encostas rochosas, ou ainda dunas arenosas. Inclui-se aí a m" \
    "aior parte dos terrenos de mangues e de pântanos e terras muito áridas, que não prestam para pastoreio.")

    par20 = doc.add_paragraph()
    run20 = par20.add_run("Hidrografia")
    run20.bold = True

    par21 = doc.add_paragraph()
    run21 = par21.add_run("Muito Bom: ")
    run21.bold = True
    run21_1 = par21.add_run("Lâmina Útil apropriada para culturas com alta demanda hídrica. " \
    "Como referência, utilizou-se a cultura de Cana de açúcar e culturas hortícolas.")

    par22 = doc.add_paragraph()
    run22 = par22.add_run("Bom: ")
    run22.bold = True
    run22_1 = par22.add_run("Lâmina Útil apropriada para culturas com demanda hídrica moderada, " \
    "como referência, foi adotada a culturas anuais (milho, feijão, trigo...)")


    par24 = doc.add_paragraph()
    run24 = par24.add_run("Normal: ")
    run24.bold = True
    run24_1 = par24.add_run("Lâmina Útil apropriada para pastagens artificiais, e " \
    "demais culturas com demanda hídrica baixa.")

    par25 = doc.add_paragraph()
    run25 = par25.add_run("Regular: ")
    run25.bold = True
    run25_1 = par25.add_run("Representa uma lâmina útil insuficiente para irrigação de culturas agrícolas. " \
    "Foi adotado em razão de representar uma vazão de curso d’água suficiente para a dessedentação dos " \
    "animais em pecuária de confinamento.")

    par26 = doc.add_paragraph()
    run26 = par26.add_run("Ruim: ")
    run26.bold = True
    run26_1 = par26.add_run("É a mera presença de um curso d’água qualquer na propriedade, " \
    "insuficiente para atender qualquer uma das situações expostas nos intervalos superiores; " \
    "entretanto, adequada para pecuária extensiva e fornecimento de água para sede do imóvel, " \
    "quando for o caso.")

    par27 = doc.add_paragraph()
    run27 = par27.add_run("Muito ruim: ")
    run27.bold = True
    run27_1 = par27.add_run("Representa, obviamente, uma propriedade com ausência de fontes próprias de água.")

    par28 = doc.add_paragraph()
    run28 = par28.add_run("Situação")
    run28.bold = True

    extra_par = doc.add_paragraph()

    par29 = doc.add_paragraph()
    run29 = par29.add_run("[INSERIR_SITUACAO_AQUI]")

    par30 = doc.add_paragraph()
    run30 = par30.add_run("Fonte")
    run30.bold = True

    par31 = doc.add_paragraph()
    run31 = par31.add_run("Considerando a tendência de supervalorização observada em imóveis " \
    "disponibilizados para comercialização, essa categoria de amostra está sujeita à influência do " \
    "denominado ""Fator Fonte"", cuja literatura aponta como sendo, em média, 10%. Dessa forma, no " \
    "processamento das amostras relacionadas a imóveis em oferta, deve-se aplicar um deságio de 10% sobre " \
    "o valor da terra nua ou do terreno, calculado com base na unidade de área correspondente.")

    par32 = doc.add_paragraph()
    run32 = par32.add_run("Amostras de Mercado")
    run32.bold = True

    for i in range (1, 7):
        par33 = doc.add_paragraph()
        run33 = par33.add_run(f"[INSERIR_TABELA_AMOSTRAL_{i:02d}]")

    par34 = doc.add_paragraph()
    run34 = par34.add_run("Quadro de amostras")
    run34.bold = True
    
    par35 = doc.add_paragraph()
    run35 = par35.add_run("Para a adequação e depuração dos dados coletados, foi empregado " \
    "o ""Tratamento por Fatores de Homogeneização"". Esse procedimento, aplicável ao método " \
    "comparativo de dados de mercado, pressupõe a existência de relações fixas entre atributos " \
    "específicos e seus respectivos preços. Para viabilizar essa abordagem, são utilizados fatores " \
    "de homogeneização calculados conforme a norma ABNT NBR 14.653-3:2019, item 7.7.2.1, garantindo " \
    "que tais fatores expressem, de maneira relativa, o comportamento do mercado dentro de um " \
    "determinado contexto espacial e temporal.")

    par36 = doc.add_paragraph()
    run36 = par36.add_run("[INSERIR_QUADRO_AQUI]")

    par37 = doc.add_paragraph()
    run37 = par37.add_run("Quadro de amostras")
    run37.bold = True

    par38 = doc.add_paragraph()
    run38 = par38.add_run("[INSERIR_HOMOG_AQUI]")

    par39 = doc.add_paragraph()
    run39 = par39.add_run("Saneamento de Amostras")
    run39.bold = True

    par40 = doc.add_paragraph()
    run40 = par40.add_run("A depuração e análise dos resultados obtidos foram realizadas considerando " \
    "uma faixa de 30% em torno da média, com a exclusão dos elementos discrepantes. O intervalo de " \
    "confiança foi determinado com base na distribuição t de Student, em conformidade com as Normas da " \
    "ABNT, garantindo um nível mínimo de certeza de 80%. Por meio da aplicação de métodos estatísticos " \
    "descritivos à amostra, conforme detalhado na memória de cálculo do item 8.2, foram identificados os " \
    "seguintes valores:")

    par41 = doc.add_paragraph()
    run41 = par41.add_run("[INSERIR_SANEAMENTO_AQUI]")

    par42 = doc.add_paragraph()
    run42 = par42.add_run("Valor de Liquidação Forçada")
    run42.bold = True

    par43 = doc.add_paragraph()
    run43 = par43.add_run("O princípio da prudência é um elemento fundamental nas avaliações realizadas " \
    "para fins de garantia, abrangendo tanto o valor de mercado quanto o valor de liquidação forçada. " \
    "A definição de valor de liquidação forçada, conforme estabelecido pela norma ABNT NBR 14.653-1, pode " \
    "ser descrita da seguinte forma:\n" \
    "Valor de liquidação forçada: trata-se da estimativa de um bem em uma situação de venda compulsória ou " \
    "dentro de um período inferior ao convencionalmente observado no mercado.\n"\
    "A determinação desse valor é realizada por meio de uma função financeira, na qual as variáveis essenciais " \
    "são o valor do imóvel, o prazo de comercialização e as taxas de juros vigentes. Esses fatores representam o " \
    "custo de oportunidade associado à necessidade de uma venda acelerada do ativo.\n"
    "O coeficiente aplicado ao valor de mercado obtido é calculado utilizando a seguinte fórmula:\n")

    par44 = doc.add_paragraph()
    run44 = par44.add_run("VP=VM×(1 - i)n")
    run44.alignment = WD_ALIGN_PARAGRAPH.CENTER

    par45 = doc.add_paragraph()
    run45 = par45.add_run("Onde:\n" \
    "•	VP = Valor de Liquidação Forçada (R$)\n" \
    "•	VM = Valor de Mercado (R$)\n"\
    "•	i = Taxa de Desconto Adotada (%)\n"\
    "•	n = Prazo de comercialização (meses)\n" \
    "Essa metodologia permite ajustar o valor do imóvel considerando o impacto das condições de venda " \
    "acelerada sobre o preço final de transação.")
    
    par46 = doc.add_paragraph()
    run46 = par46.add_run("[INSERIR_LIQUIDACAO_AQUI]")

    par47 = doc.add_paragraph()
    run47 = par47.add_run("Referencias")
    run47.bold = True

    par48 = doc.add_paragraph()
    run48 = par48.add_run("ABNT – Associação Brasileira de Normas Técnicas. NBR nº 14.653:1 (2019) e nº 14.653:3 (2019).\n"\
    "Abunahman, Sérgio Antonio. Engenharia Legal e de Avaliações. Pini- 4ª ed., 2000.\n"\
    "Alves, C. S. Método Prático de Determinação de Percentual de Servidão para Faixa e Áreas Remanescentes. Revista de Avaliações e Perícias. IBAPE-RS, 2002.\n"\
    "Arantes, Carlos Augusto. Depreciação de Área remanescente por Apossamento Administrativo. Fortaleza: XIII COBREAP, 2006.\n"\
    "Arantes, Carlos Augusto; Saldanha, Marcelo Suarez. Avaliações de Imóveis Rurais. São Paulo: Leud, 2009.\n"\
    "DESLANDES, C.A. Avaliações de Imóveis Rurais. Editora Aprenda Fácil. Viçosa/MG, 2002.\n"\
    "Estado do Rio Grande do Sul. Modelo Rural – Requisitos Mínimos para Laudo de Avaliação. Governo do Estado do Rio Grande do Sul, 2020.\n"\
    "Hantzis, et al. Indemnizaciones por Concepto de Imposición de Servidumbres de Gasoducto. CBAP, 2000.\n"\
    "LIMA, M. R. C. Avaliação de Propriedades Rurais. Editora Leud: São Paulo/SP, 2011.\n"\
    "Manual Brasileiro para Levantamento da Capacidade de Uso da Terra (ETA – Escritório Técnico de Agricultura Brasil – Estados Unidos) III aproximação.\n"\
    "Manual para Classificação da Capacidade de Uso das Terras para fins de Avaliação de Imóveis Rurais – 1º aproximação/CESP.\n"\
    "PELLEGRINO, J. C. Engenharia de Avaliações. São Paulo: Editora Pini; 1974.\n"\
    "Resolução n.º 342/90 do CONFEA, que dispõe sobre a responsabilidade técnica do engenheiro agrônomo.\n"\
    "Sindicato Nacional")

    
    return doc

