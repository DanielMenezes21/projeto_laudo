from win32com.client import Dispatch, constants
import os

def inserir_tabela_excel_no_word(docx_path, excel_path, aba="AMOSTRAS", largura_maxima_cm=16):
    """
    Insere uma tabela do Excel no Word, usando o caminho do Excel selecionado pelo usuário.
    """
    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    if wb is None:
        excel.Quit()
        raise Exception(f"Não foi possível abrir o arquivo Excel: {excel_path}")

    abas_disponiveis = [s.Name for s in wb.Sheets]
    if aba not in abas_disponiveis:
        wb.Close(SaveChanges=False)
        excel.Quit()
        raise Exception(f"Aba '{aba}' não encontrada no arquivo Excel: {excel_path}\nAbas disponíveis: {abas_disponiveis}")

    sheet = wb.Sheets(aba)

    last_row = sheet.Cells(sheet.Rows.Count, "B").End(xlUp).Row
    linha_final = None
    for row in range(4, last_row + 1):
        valor = str(sheet.Cells(row, 2).Value or "").strip().lower()
        if "amostral" in valor:
            linha_final = row - 2
            break
    if linha_final is None:
        wb.Close(SaveChanges=False)
        excel.Quit()
        raise Exception("Texto 'amostral' não encontrado.")

    intervalo = f"B3:I{linha_final}"
    sheet.Range(intervalo).Copy()

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    word.Selection.HomeKey(Unit=6)  
    if word.Selection.Find.Execute("[INSERIR_TABELA_AQUI]"):
        word.Selection.TypeBackspace()  
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()

def inserir_tabela_benfeitoria_no_word(docx_path, excel_path, aba="FATORES", largura_maxima_cm=16):
    """
    Insere uma tabela do Excel no Word, usando o caminho do Excel selecionado pelo usuário.
    """
    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    last_row = sheet.Cells(sheet.Rows.Count, "D").End(xlUp).Row

    intervalo = f"D21:E27"
    sheet.Range(intervalo).Copy()

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    word.Selection.HomeKey(Unit=6)  
    if word.Selection.Find.Execute("[INSERIR_BENFEITORIA_AQUI]"):
        word.Selection.TypeBackspace() 
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()

def inserir_tabela_depreciacao_no_word(docx_path, excel_path, aba="FATORES", largura_maxima_cm=16):
    """
    Insere uma tabela do Excel no Word, usando o caminho do Excel selecionado pelo usuário.
    """
    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    last_row = sheet.Cells(sheet.Rows.Count, "H").End(xlUp).Row

    intervalo = f"H21:L27"
    sheet.Range(intervalo).Copy()

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    word.Selection.HomeKey(Unit=6)  
    if word.Selection.Find.Execute("[INSERIR_DEPRECIACAO_AQUI]"):
        word.Selection.TypeBackspace() 
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()

def inserir_tabela_classe_no_word(docx_path, excel_path, aba="UTIL_LAUDO", largura_maxima_cm=16):
    """
    Insere uma tabela do Excel no Word, usando o caminho do Excel selecionado pelo usuário.
    """
    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    last_row = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).Row

    intervalo = f"A1:C17"
    sheet.Range(intervalo).Copy()

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    word.Selection.HomeKey(Unit=6)  
    if word.Selection.Find.Execute("[INSERIR_CLASSE_AQUI]"):
        word.Selection.TypeBackspace()  
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()

def inserir_tabela_situacao_no_word(docx_path, excel_path, aba="FATORES", largura_maxima_cm=16):
    """
    Insere uma tabela do Excel no Word, usando o caminho do Excel selecionado pelo usuário.
    """
    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    last_row = sheet.Cells(sheet.Rows.Count, "D").End(xlUp).Row

    intervalo = f"D32:G39"
    sheet.Range(intervalo).Copy()

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    word.Selection.HomeKey(Unit=6)  
    if word.Selection.Find.Execute("[INSERIR_SITUACAO_AQUI]"):
        word.Selection.TypeBackspace()  
        word.Selection.Paste()

        table = doc.Tables(doc.Tables.Count)
        usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

        max_width = largura_maxima_cm * 28.35
        final_width = min(usable_width, max_width)

        table.PreferredWidthType = 1
        table.PreferredWidth = final_width
        print("sucesso")
    else:
        print("erro")

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()

def inserir_tabelas_amostras_auto(
    docx_path,
    excel_path,
    aba="AMOSTRAS",
    titulo_base="Dado Amostral",
    marcador_base="[INSERIR_TABELA_AMOSTRAL_{:02d}]",
    largura_maxima_cm=16
):
    """
    Procura todas as tabelas no Excel cujo título começa com 'Dados Amostrais',
    copia cada uma e insere no Word em marcadores sequenciais.
    """

    xlUp = -4162
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_path)
    sheet = wb.Sheets(aba)

    linhas_titulo = []
    last_row = sheet.Cells(sheet.Rows.Count, 2).End(xlUp).Row
    for row in range(1, last_row + 1):
        valor = str(sheet.Cells(row, 2).Value or "").strip()
        if valor.startswith(titulo_base):
            linhas_titulo.append(row)

    intervalos = []
    for i, linha_inicio in enumerate(linhas_titulo):
        linha_fim = (linhas_titulo[i + 1] - 2) if i + 1 < len(linhas_titulo) else last_row
        intervalo = f"B{linha_inicio}:I{linha_fim}"
        intervalos.append(intervalo)

    word = Dispatch("Word.Application")
    try:
        word.Visible = False
    except AttributeError:
        pass
    doc = word.Documents.Open(docx_path)

    for idx, intervalo in enumerate(intervalos):
        sheet.Range(intervalo).Copy()
        marcador = marcador_base.format(idx + 1)
        word.Selection.HomeKey(Unit=6)
        if word.Selection.Find.Execute(marcador):
            word.Selection.TypeBackspace()
            word.Selection.Paste()
            table = doc.Tables(doc.Tables.Count)
            usable_width = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
            max_width = largura_maxima_cm * 28.35
            final_width = min(usable_width, max_width)
            table.PreferredWidthType = 1
            table.PreferredWidth = final_width

    doc.Save()
    doc.Close()
    wb.Close(SaveChanges=False)
    excel.Quit()