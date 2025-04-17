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
from app.screen4_dadosp.dados_function import *

def go_back(self, *args):
    self.manager.current_screen.manager.current = "leitor"

def go_next(self):
    if self.caminho_car and self.caminho_cit:
        dados_screen = self.manager.get_screen("dados")
        dados_screen.preencher_com_dados(self.caminho_car, self.caminho_cit)
        self.manager.current = "dados"
    else:
        print("Ambos os PDFs precisam ser selecionados.")

def selecionar_pdf_car(self, caminho_pdf):
    self.caminho_car = caminho_pdf
    adicionar_item(self,caminho_pdf, "car")  

def selecionar_pdf_cit(self, caminho_pdf):
    self.caminho_cit = caminho_pdf
    adicionar_item(self,caminho_pdf, "cit")  

def carregar_estrutura_pasta(self, caminho_pasta, tipo):
    """Função para carregar arquivos PDF dentro de uma pasta"""
    print(f"Carregando arquivos na pasta: {caminho_pasta}")
    for nome in sorted(os.listdir(caminho_pasta)):
        caminho_completo = os.path.join(caminho_pasta, nome)
        if os.path.isfile(caminho_completo) and nome.lower().endswith(".pdf"):
            adicionar_item(self, caminho_completo, tipo)

def adicionar_item(self, caminho_pdf, tipo):
    texto = f"Arquivo: {os.path.basename(caminho_pdf)}"
    if tipo == "car":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_car(self, f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CAR] {texto}"))
        self.file_list.add_widget(item)
    elif tipo == "cit":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_cit(self,f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CIT] {texto}"))
        self.file_list2.add_widget(item)

def carregar_estrutura(self, tipo):
    print(f"Carregando estrutura para: {tipo}")
    if not os.path.exists(self.current_path):
        print(f"Caminho não encontrado: {self.current_path}")
        self.mostrar_erro(f"Caminho não encontrado: {self.current_path}")
        return

    if tipo == "car":
        self.file_list.clear_widgets()
    elif tipo == "cit":
        self.file_list2.clear_widgets()

    print(f"Listando arquivos no diretório: {self.current_path}")
    arquivos_adicionados = 0

    for nome in sorted(os.listdir(self.current_path)):
        caminho_completo = os.path.join(self.current_path, nome)

        if os.path.isdir(caminho_completo):
            carregar_estrutura_pasta(self,caminho_completo, tipo)
        elif os.path.isfile(caminho_completo) and nome.lower().endswith(".pdf"):
            adicionar_item(self,caminho_completo, tipo)
            arquivos_adicionados += 1

    if arquivos_adicionados == 0:
        print(f"Nenhum arquivo PDF encontrado no diretório {self.current_path}")


def load_directory(self,tipo, *args):
    print(f"Carregando diretório: {self.current_path}")
    self.pastas_abertas = set() 
    carregar_estrutura(self, tipo)

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
    modelo_path = os.path.join(os.getcwd(), "models","MODELO_LAUDO.docx")
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
                print(f"Imagem inserida em {placeholder}")
                return True
        return False

    try:
        imagens = extrair_paginas_como_imagens(caminho_pdf)  # Aqui o processo de extração das imagens
        if not imagens:
            print("Nenhuma imagem extraída do PDF.")
            return

        print(f"Imagens extraídas: {imagens}")  # Log para verificar se as imagens foram extraídas corretamente

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
