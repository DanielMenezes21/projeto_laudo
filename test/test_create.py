from test_page2 import *
from test_image import *
from test_page3 import *
from test_page4 import *
from test_page5 import *
from test_page6 import *
import os

def create_document(valor_texto, imagem_path='captura_teste.png', celula_verde=False, imagem_acesso=None):
    """Função principal que cria toda a ficha cadastral"""
    doc = configurar_documento()
    doc.add_page_break()
    doc = criar_titulo(doc)
    doc = adicionar_linha_fina(doc)
    doc = criar_secao_valor(doc, valor_texto)
    doc = adicionar_linha_fina(doc)
    doc = criar_secao_identificacao(doc)
    doc = adicionar_linha_fina(doc)
    doc = criar_secao_croqui(doc, imagem_path)
    doc = adicionar_linha_fina(doc)
    doc, tabela = geometria_terreno(doc)
    if celula_verde:
        colorir_celula(tabela.rows[1].cells[1], "009933")
    else:
        colorir_celula(tabela.rows[1].cells[1], "FFFFFF")
    doc = adicionar_linha_fina(doc)
    doc = criar_secao_caracteristicas(doc)
    doc = titulo(doc)
    doc = table_geo(doc)
    doc = adicionar_linha_fina(doc)
    doc = tabela_bioma(doc)
    doc = adicionar_linha_fina(doc)
    doc = area_APA(doc)
    doc = adicionar_linha_fina(doc)
    doc = table_passivo_ambiental(doc)
    doc = campo_assinatura(doc)
    doc.add_page_break()
    doc = inserir_sumario(doc)
    doc.add_page_break()
    doc = adicionar_espaco(doc)
    doc = texto_solicitante(doc)
    doc = adicionar_espaco(doc)
    doc = texto_objetivo(doc)
    doc = adicionar_espaco(doc)
    doc = texto_finalidade(doc)
    doc = adicionar_espaco(doc)
    doc = texto_proprietario(doc)
    doc = adicionar_espaco(doc)
    doc = texto_ressalvas(doc)
    doc.add_page_break()
    doc = title_imovel(doc)
    doc = localizacao(doc)
    doc = acesso(doc,imagem_acesso)
    doc.add_page_break()

    doc.save('ficha_cadastral_final.docx')
    return doc

def gerar_documento_completo(valor, texto_capa,celula_verde):
    """Gera o docx com texto, insere a imagem de capa atrás do texto na primeira página,
    a marca d'água em todas as páginas, a imagem na última página e uma caixa de texto opcional na capa."""
    doc = create_document(valor, celula_verde=celula_verde)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docx_path = os.path.join(base_dir, "ficha_cadastral_final.docx")
    img_marca = os.path.join(base_dir, "models", "RODAPE.png")
    img_capa = os.path.join(base_dir, "models", "capa_do_laudo.png")
    img_fim = os.path.join(base_dir, "models", "final.png")

    inserir_imagem_capa_atras_texto(docx_path, img_capa)
    imagens_fundo(docx_path, img_marca)
    inserir_imagem_ultima_pagina(docx_path, img_fim)
    inserir_caixa_texto_primeira_pagina(docx_path, texto_capa)