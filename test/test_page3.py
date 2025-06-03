import win32com.client
import os
from test_page2 import pagina_2
from test_page1 import inserir_caixa_texto_primeira_pagina, inserir_imagem_capa_atras_texto

def inserir_imagem_ultima_pagina(docx_path, img_fim):
    """Insere uma imagem atrás do texto na última página do documento."""
    wdWrapBehind = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(docx_path)

    sel = word.Selection
    sel.EndKey(Unit=6)  # Vai para o final do documento

    shape = doc.Shapes.AddPicture(
        FileName=img_fim,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=0,
        Top=0,
        Width=doc.PageSetup.PageWidth,
        Height=doc.PageSetup.PageHeight
    )
    shape.WrapFormat.Type = wdWrapBehind
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    shape.Left = 0
    shape.Top = 0

    doc.SaveAs(docx_path)
    doc.Close()
    word.Quit()

def imagens_fundo(docx_path, img_marca):
    """Insere uma marca d'água em todas as páginas."""
    wdHeaderFooterPrimary = 1
    wdWrapBehind = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)

    for section in doc.Sections:
        header = section.Headers(wdHeaderFooterPrimary)
        shape = header.Shapes.AddPicture(
            FileName=img_marca,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=word.ActiveDocument.PageSetup.PageWidth,
            Height=word.ActiveDocument.PageSetup.PageHeight
        )
        shape.WrapFormat.Type = wdWrapBehind
        shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
        shape.Left = 0
        shape.Top = 0

    doc.SaveAs(docx_path)
    doc.Close()
    word.Quit()

def gerar_documento_completo(valor, texto_capa):
    """Gera o docx com texto, insere a imagem de capa atrás do texto na primeira página,
    a marca d'água em todas as páginas, a imagem na última página e uma caixa de texto opcional na capa."""
    pagina_2(valor)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docx_path = os.path.join(base_dir, "ficha_cadastral_final.docx")
    img_marca = os.path.join(base_dir, "models", "RODAPE.png")
    img_capa = os.path.join(base_dir, "models", "capa_do_laudo.png")
    img_fim = os.path.join(base_dir, "models", "final.png")

    inserir_imagem_capa_atras_texto(docx_path, img_capa)
    imagens_fundo(docx_path, img_marca)
    inserir_imagem_ultima_pagina(docx_path, img_fim)
    inserir_caixa_texto_primeira_pagina(docx_path, texto_capa)