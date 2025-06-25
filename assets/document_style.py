import win32com.client
import os

def inserir_imagem_ultima_pagina(docx_path, img_fim):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        sel = word.Selection
        sel.EndKey(Unit=6)  # Vai para o fim do documento

        shape = doc.Shapes.AddPicture(
            FileName=os.path.abspath(img_fim),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=doc.PageSetup.PageWidth,
            Height=doc.PageSetup.PageHeight
        )
        shape.WrapFormat.Type = 3  
        shape.RelativeHorizontalPosition = 1
        shape.RelativeVerticalPosition = 1
        shape.Left = 0
        shape.Top = 0

        doc.Save()
    except Exception as e:
        print(f"❌ Erro ao inserir imagem final: {e}")
    finally:
        doc.Close(False)
        word.Quit()


def imagens_fundo(docx_path, img_marca):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        for section in doc.Sections:
            header = section.Headers(1)  
            shape = header.Shapes.AddPicture(
                FileName=os.path.abspath(img_marca),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=0,
                Top=0,
                Width=doc.PageSetup.PageWidth,
                Height=doc.PageSetup.PageHeight
            )
            shape.WrapFormat.Type = 3
            shape.RelativeHorizontalPosition = 1
            shape.RelativeVerticalPosition = 1
            shape.Left = 0
            shape.Top = 0

        doc.Save()
    except Exception as e:
        print(f"❌ Erro ao inserir marca d'água: {e}")
    finally:
        doc.Close(False)
        word.Quit()
