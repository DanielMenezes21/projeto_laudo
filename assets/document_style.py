import win32com.client
import os

def inserir_imagem_ultima_pagina(docx_path, img_fim):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))

        largura_a4_pt = 21.0 * 28.35
        altura_a4_pt = 29.7 * 28.35

        end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

        shape = doc.Shapes.AddPicture(
            FileName=os.path.abspath(img_fim),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=largura_a4_pt,
            Height=altura_a4_pt,
            Anchor=end_range
        )

        shape.WrapFormat.Type = 3  
        shape.LockAspectRatio = False
        shape.RelativeHorizontalPosition = 1  
        shape.RelativeVerticalPosition = 1    
        shape.Left = 0
        shape.Top = 0
        shape.ZOrder(4)

        doc.Save()
        print("✅ Imagem final inserida na última página (A4 retrato).")

    except Exception as e:
        print(f"❌ Erro ao inserir imagem final: {e}")
    finally:
        doc.Close(False)
        #word.Quit()

def imagens_fundo(docx_path, img_marca, ignorar_secao=None):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))

        largura_a4_pt = 21.6 * 28.35  
        altura_a4_pt = 29.7 * 28.35   

        for i, section in enumerate(doc.Sections, start=1):
            if ignorar_secao and i == ignorar_secao:
                continue

            if section.PageSetup.Orientation == 0:  
                header = section.Headers(1)

                shape = header.Shapes.AddPicture(
                    FileName=os.path.abspath(img_marca),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=0,
                    Top=0,
                    Width=largura_a4_pt,
                    Height=altura_a4_pt
                )

                shape.LockAspectRatio = True
                shape.WrapFormat.Type = 3  
                shape.RelativeHorizontalPosition = 1 
                shape.RelativeVerticalPosition = 1    
                shape.Left = 0
                shape.Top = 0
                shape.ZOrder(0)

        doc.Save()
        print("✅ Marca d'água aplicada em páginas retrato com tamanho fixo.")
    except Exception as e:
        print(f"❌ Erro ao inserir marca d'água: {e}")
    finally:
        doc.Close(False)
        #word.Quit()
