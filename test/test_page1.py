import win32com.client
import os

def inserir_imagem_capa_atras_texto(docx_path, img_capa):
    """Insere uma imagem atrás do texto apenas na primeira página."""
    wdWrapBehind = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)

    shape = doc.Shapes.AddPicture(
        FileName=img_capa,
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

def inserir_caixa_texto_primeira_pagina(docx_path, texto, width=400, height=100):
    """Insere uma caixa de texto (TextBox) centralizada à direita na primeira página."""
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)

    page_width = doc.PageSetup.PageWidth
    page_height = doc.PageSetup.PageHeight

    right_margin = 57
    left = page_width - width - right_margin
    top = (page_height - height) // 2

    shape = doc.Shapes.AddTextbox(
        Orientation=1, 
        Left=left,
        Top=top,
        Width=width,
        Height=height
    )
    shape.Line.Visible = False
    shape.TextFrame.TextRange.Text = texto
    shape.TextFrame.TextRange.Font.Size = 22
    shape.TextFrame.TextRange.Font.Bold = True
    shape.TextFrame.TextRange.Font.Color = 16777215  # Branco
    shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1  # Center

    doc.SaveAs(docx_path)
    doc.Close()
    word.Quit()