import fitz  # PyMuPDF
import re
import os
from docx import Document

def extrair_dados_pdf(caminho_pdf):
    """
    Abre o PDF, varre todas as páginas em busca de Nome e CPF,
    e retorna (nome, cpf).
    """
    nome = ""
    cpf = ""
    try:
        pdf = fitz.open(caminho_pdf)
        for pagina in pdf:
            texto = pagina.get_text()
            # Ajuste as regex se o padrão for diferente
            m_nome = re.search(r"\bNome[:\-]?\s*(.+)", texto, re.IGNORECASE)
            m_cpf  = re.search(r"\bCPF[:\-]?\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})", texto)
            if m_nome:
                nome = m_nome.group(1).strip()
            if m_cpf:
                cpf = m_cpf.group(1).strip()
            if nome and cpf:
                break
        pdf.close()
    except Exception as e:
        print(f"❌ Erro ao extrair dados do PDF: {e}")
    return nome or "", cpf or ""

def carregar_pdf_dados(self, caminho_car, caminho_cit):
    self.caminho_car = caminho_car
    self.caminho_cit = caminho_cit

    nome, cpf = extrair_dados_pdf(caminho_car)  
    self.proponente.text = nome
    self.cpf.text = cpf

    placeholder_car = "#SUBSTITUIR_CAR"  
    self.inserir_pdf_no_word(self.caminho_car, placeholder_car)

def receber_dados(self):
    nome = self.proponente.text
    cpf = self.cpf.text

    # Caminho do modelo Word
    modelo_path = os.path.join(os.getcwd(), "models", "MODELO_LAUDO.docx")
    doc = Document(modelo_path)

    # Substituir placeholders no Word
    for par in doc.paragraphs:
        for run in par.runs:
            run.text = run.text.replace("#PROPONENTE", nome).replace("#CPF_PROPONENTE", cpf)

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for par in celula.paragraphs:
                    for run in par.runs:
                        run.text = run.text.replace("#PROPONENTE", nome).replace("#CPF_PROPONENTE", cpf)

    nome_limpo = re.sub(r"[^\w\s-]", "", nome).replace(" ", "_")
    saida_path = os.path.join(os.getcwd(), f"LAUDO_{nome_limpo}.docx")
    doc.save(saida_path)
    print(f"✅ Documento salvo em: {saida_path}")