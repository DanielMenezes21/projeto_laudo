from win32com.client import Dispatch, constants
import os

def inserir_tabela_dinamica_no_word(docx_path, excel_path=None, aba="AMOSTRAS", largura_maxima_cm=16):
    if excel_path is None:
        excel_path = os.path.abspath("models/LAUDO DE AVALIAÇÃO N° 12941255 - MAURICIO MIYASAKI.xlsx")

    # Abrir Excel
    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    # Localizar até 2 linhas antes de "amostral"
    last_row = sheet.Cells(sheet.Rows.Count, "B").End(constants.xlUp).Row
    linha_final = None
    for row in range(4, last_row + 1):
        valor = str(sheet.Cells(row, 2).Value or "").strip().lower()
        if "amostral" in valor:
            linha_final = row - 2
            break
    if linha_final is None:
        raise Exception("Texto 'amostral' não encontrado.")

    intervalo = f"B9:I{linha_final}"
    sheet.Range(intervalo).Copy()

    # Abrir Word
    word = Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(docx_path)

    # Procurar marcador no documento
    word.Selection.HomeKey(Unit=6)  # Início do documento
    if word.Selection.Find.Execute("[INSERIR_TABELA_AQUI]"):
        word.Selection.TypeBackspace()  # remove o marcador
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        # Limite em pontos (1 cm = 28.35 points)
        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()
