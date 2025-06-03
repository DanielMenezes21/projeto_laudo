import win32com.client
import os

wdHeaderFooterPrimary = 1
wdWrapBehind = 3
wdRelativeHorizontalPositionPage = 1
wdRelativeVerticalPositionPage = 1

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
docx_path = os.path.join(base_dir, "entrada.docx")
img_path = os.path.join(base_dir, "models", "RODAPE.png")
path_destino = os.path.join(base_dir, "saida.docx")

word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(docx_path)

for section in doc.Sections:
    header = section.Headers(wdHeaderFooterPrimary)
    shape = header.Shapes.AddPicture(
        FileName=img_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=100,  # ajuste a posição conforme necessário
        Top=100,
        Width=300,  # ajuste o tamanho conforme necessário
        Height=300
    )
    shape.WrapFormat.Type = wdWrapBehind
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    shape.PictureFormat.TransparencyColor = 16777215  # Branco como transparente (opcional)
    #shape.PictureFormat.Transparency = 0.5  # 0 = opaco, 1 = totalmente transparente (pode variar conforme versão do Word)
    shape.Left = 100
    shape.Top = 100

doc.SaveAs()
doc.Close()
word.Quit()