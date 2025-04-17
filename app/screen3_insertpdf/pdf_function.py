import os
import re
import uuid
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
from pdf2image import convert_from_path
from kivymd.uix.list import MDListItem, MDListItemHeadlineText
from kivy.uix.popup import Popup
from app.screen3_insertpdf.pdf_function2 import extrair_paginas_como_imagens


def go_back(self, *args):
    self.manager.current_screen.manager.current = "leitor"

def go_next(self, *args):
    if hasattr(self, "caminho_car") and hasattr(self, "caminho_cit"):
        dados_screen = self.manager.get_screen("dados")
        dados_screen.carregar_pdf_dados(self.caminho_car, self.caminho_cit)
        self.manager.current = "dados"
    else:
        print("Erro", "Selecione os dois PDFs (CAR e CIT) antes de prosseguir.")

# Seleção do PDF CAR
def selecionar_pdf_car(self, caminho_pdf):
    self.caminho_car = caminho_pdf
    self.adicionar_item(caminho_pdf, "car")  # Adiciona o arquivo CAR à lista

# Seleção do PDF CIT
def selecionar_pdf_cit(self, caminho_pdf):
    self.caminho_cit = caminho_pdf
    self.adicionar_item(caminho_pdf, "cit")  # Adiciona o arquivo CIT à lista

def adicionar_item(self, caminho_pdf, tipo):
    texto = f"Arquivo: {os.path.basename(caminho_pdf)}"
    if tipo == "car":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: self.selecionar_pdf_car(f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CAR] {texto}"))
        self.file_list.add_widget(item)
    elif tipo == "cit":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: self.selecionar_pdf_cit(f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CIT] {texto}"))
        self.file_list2.add_widget(item)

# Função para carregar os arquivos selecionados
def carregar_estrutura(self, tipo):
    if tipo == "car":
        self.file_list.clear_widgets()
    elif tipo == "cit":
        self.file_list2.clear_widgets()

    def listar_pdf_e_pastas(caminho, nivel=0):
        prefixo = "    " * nivel
        for nome in sorted(os.listdir(caminho)):
            caminho_completo = os.path.join(caminho, nome)
            if os.path.isdir(caminho_completo):
                # Função para adicionar pastas
                pass
            elif nome.lower().endswith(".pdf"):
                adicionar_item(nome, caminho_completo, prefixo=prefixo)

    listar_pdf_e_pastas(self.current_path)
    
def load_directory(self, *args):
    self.pastas_abertas = set()  # Começa do zero
    carregar_estrutura(self)

def entrar_em_pasta(self, caminho_pasta, tipo):
    if caminho_pasta in self.pastas_abertas:
        self.pastas_abertas.remove(caminho_pasta)
    else:
        self.pastas_abertas.add(caminho_pasta)
    carregar_estrutura(self, tipo)

def go_up(self, *args):
    self.current_path = os.path.dirname(self.current_path)
    if not os.path.exists(self.current_path):
        print(f"Diretório não encontrado: {self.current_path}")
        return
    carregar_estrutura(self)

def initialize_word(self):
    modelo_path = os.path.join(os.getcwd(), "models\\MODELO_LAUDO.docx")
    if not os.path.exists(modelo_path):
        print(f"Modelo não encontrado: {modelo_path}")
        return
    self.doc = Document(modelo_path)

def inserir_pdf_no_word(self, caminho_pdf, placeholder):
    if not hasattr(self, "doc"):
        print("Documento Word não inicializado!")
        return

    def substituir_em_paragrafos(paragrafos):
        for par in paragrafos:
            texto_completo = ''.join(run.text for run in par.runs)
            if placeholder in texto_completo:
                for run in par.runs:
                    run.text = ""
                for imagem_path in imagens:
                    novo_run = par.add_run()
                    novo_run.add_picture(imagem_path, width=Inches(6))
                return True
        return False

    try:
        imagens = extrair_paginas_como_imagens(caminho_pdf)
        if not imagens:
            print("Nenhuma imagem extraída do PDF.")
            return

        encontrado = substituir_em_paragrafos(self.doc.paragraphs)

        if not encontrado:
            for tabela in self.doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        if substituir_em_paragrafos(celula.paragraphs):
                            encontrado = True
                            break
                    if encontrado:
                        break
                if encontrado:
                    break

        if encontrado:
            print(f"PDF '{caminho_pdf}' inserido com sucesso no local '{placeholder}'.")
            if not hasattr(self, "pdfs_inseridos"):
                self.pdfs_inseridos = set()
            self.pdfs_inseridos.add(placeholder)
        else:
            print(f"❌ Placeholder '{placeholder}' não encontrado no documento.")

    except Exception as e:
        print(f"Erro ao inserir PDF no Word: {e}")