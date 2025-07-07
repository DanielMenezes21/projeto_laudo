import win32com.client
import os

def inserir_caixa_texto_primeira_pagina(docx_path, texto, width=400, height=100, substituicoes=None):
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        if substituicoes:
            for chave, valor in substituicoes.items():
                texto = texto.replace(chave, str(valor))

        page_width = doc.PageSetup.PageWidth
        page_height = doc.PageSetup.PageHeight

        # Alinhar totalmente à direita, ignorando margens
        left = page_width - width  # 0 = esquerda, page_width-width = direita total
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

        shape.TextFrame.TextRange.Font.Color = 16777215  

        shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  

        doc.Save()
        print("✅ Caixa de texto inserida na capa!")
    except Exception as e:
        print(f"❌ Erro ao inserir caixa de texto: {e}")
    finally:
        try:
            doc.Close(False)
            word.Quit()
        except Exception as close_err:
            print(f"⚠️ Erro ao tentar fechar o Word: {close_err}")

def inserir_imagem_capa_atras_texto(docx_path, img_capa):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        shape = doc.Shapes.AddPicture(
            FileName=os.path.abspath(img_capa),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=doc.PageSetup.PageWidth,
            Height=doc.PageSetup.PageHeight
        )
        shape.ZOrder(4) 

        shape.WrapFormat.Type = 3
        shape.RelativeHorizontalPosition = 1
        shape.RelativeVerticalPosition = 1
        shape.Left = 0
        shape.Top = 0

        doc.Save()
    except Exception as e:
        print(f"❌ Erro ao inserir imagem de capa: {e}")
    finally:
        try:
            doc.Close(False)
            word.Quit()
        except Exception as close_err:
            print(f"⚠️ Erro ao tentar fechar o Word: {close_err}")