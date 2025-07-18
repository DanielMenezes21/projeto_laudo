import os
import re
from docx import Document
from kivymd.uix.list import MDListItem, MDListItemHeadlineText
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.metrics import dp
from kivymd.uix.filemanager import MDFileManager
from app.screen3_dadosp.dados_function import *
from modules.resource_path import resource_path
from docx.shared import Pt

def go_back(self, *args):
    self.manager.current_screen.manager.current = "territorio"

def selecionar_pdf_car(self, caminho_pdf):
    self.caminho_car = caminho_pdf
    adicionar_item(self, caminho_pdf, "car")  

def selecionar_pdf_cit(self, caminho_pdf):
    self.caminho_cit = caminho_pdf
    adicionar_item(self, caminho_pdf, "cit")  

def carregar_estrutura_pasta(self, caminho_pasta, tipo):
    """Função para carregar arquivos PDF dentro de uma pasta"""
    for nome in sorted(os.listdir(caminho_pasta)):
        caminho_completo = os.path.join(caminho_pasta, nome)
        if os.path.isdir(caminho_completo):
            match = re.match(r"PROCESSO\s*Nº\s*(\d+)", nome)
            if match:
                numero_processo = match.group(1)
                self.numero_processo = numero_processo  
                if numero_processo == "":
                    MDSnackbar(
                        MDSnackbarText(text="Número do processo não encontrado!"),
                        y=dp(24)
                    ).open()
                    return
                carregar_estrutura_pasta(self, caminho_completo, tipo)
        elif os.path.isfile(caminho_completo) and nome.lower().endswith(".pdf"):
            adicionar_item(self, caminho_completo, tipo)

def abrir_gerenciador(self, tipo):
    """
    Abre o MDFileManager configurado para exibir apenas pastas e arquivos PDF.
    """
    self.tipo_selecionado = tipo
    if not hasattr(self, "file_manager") or self.file_manager is None:
        self.file_manager = MDFileManager(
            exit_manager=lambda *args: exit_file_manager(self, *args),
            select_path=lambda path: select_path(self, path),
            preview=False,
            search='all'
        )
    self.file_manager.show(self.current_path if hasattr(self, "current_path") else os.path.expanduser("~"))
    self.manager_open = True

def exit_file_manager(self, *args):
    """
    Fecha o MDFileManager.
    """
    self.manager_open = False
    if hasattr(self, "file_manager") and self.file_manager:
        self.file_manager.close()

def select_path(self, path):
    """
    Função chamada ao selecionar um arquivo ou pasta no MDFileManager.
    """
    exit_file_manager(self)
    if os.path.isdir(path):
        self.current_path = path
        abrir_gerenciador(self, self.tipo_selecionado)
    elif os.path.isfile(path) and path.lower().endswith(".pdf"):
        if self.tipo_selecionado == "car":
            self.caminho_car = path
        elif self.tipo_selecionado == "cit":
            self.caminho_cit = path
        adicionar_item(self, path, self.tipo_selecionado)
    else:
        MDSnackbar(
            MDSnackbarText(text="Seleção inválida. Escolha um arquivo PDF."),
            y=dp(24)
        ).open()

def adicionar_item(self, caminho_pdf, tipo):
    """
    Adiciona o arquivo PDF selecionado à lista correspondente.
    """
    texto = f"Arquivo: {os.path.basename(caminho_pdf)}"
    if tipo == "car":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_car(self, f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CAR] {texto}"))
        self.file_list.add_widget(item)
    elif tipo == "cit":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_cit(self, f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CIT] {texto}"))
        self.file_list2.add_widget(item)

def carregar_estrutura(self, tipo):
    print(f"Carregando estrutura para: {tipo}")
    if not os.path.exists(self.current_path):
        MDSnackbar(
            MDSnackbarText(text="Caminho não encontrado!"),
            y=dp(24)
        ).open()
        return

    if tipo == "car":
        self.file_list.clear_widgets()
    elif tipo == "cit":
        self.file_list2.clear_widgets()

    print(f"Listando arquivos no diretório: {self.current_path}")
    arquivos_adicionados = 0
    for arquivo in os.listdir(self.current_path):
        if arquivo.lower().endswith(".pdf"):
            arquivos_adicionados += 1
               
    if arquivos_adicionados == 0:
        MDSnackbar(
            MDSnackbarText(text="Nenhum arquivo PDF encontrado na pasta."),
            y=dp(24)
        ).open()

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
        MDSnackbar(
            MDSnackbarText(text="Caminho não encontrado!"),
            y=dp(24)
        ).open()
        return
    carregar_estrutura(self)

def initialize_word(self):
    modelo_path = resource_path(os.path.join("models", "MODELO_LAUDO.docx"))
    if not os.path.exists(modelo_path):
        MDSnackbar(
            MDSnackbarText(text="Modelo de documento não encontrado!"),
            y=dp(24)
        ).open()
        return
    self.doc = Document(modelo_path)

