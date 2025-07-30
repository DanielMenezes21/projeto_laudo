from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.dialog import MDDialog, MDDialogContentContainer, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.list import MDList, MDListItem, MDListItemTrailingCheckbox, MDListItemLeadingIcon, MDListItemHeadlineText
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.selectioncontrol import MDCheckbox
from kivy.metrics import dp
from kivymd.uix.label import MDLabel, MDIcon
from docx import Document
from modules.pesquisa import buscar_descricao_cidade
from modules.resource_path import resource_path
from app.screen4_territorio.solotable import extrair_textos_e_tabelas
import os

def go_back(self):
    self.manager.current_screen.manager.current = "dados"

def go_next1(self):
    campos = {
        "descricao_cidade": self.descricao_cidade.text,
        "regiao_imovel": self.regiao_imovel.text,
        "declividade": self.declividade_text.text,
        "hidrografia": self.hidrografia_text.text,
        "resumo_solo": self.resumo_solo.text,
        "texto_solos": self.texto_solos.text,
        "data_vistoria": self.data_vistoria_text.text,
        "rotas": self.rotas_text.text,
    }

    campos_vazios = [nome for nome, valor in campos.items() if not valor.strip()]

    if campos_vazios:
        texto_erro = "Os seguintes campos estão vazios:\n" + "\n".join(f"- {campo}" for campo in campos_vazios)
        dialog = MDDialog(
            MDDialogHeadlineText(text="Campos obrigatórios não preenchidos"),
            MDDialogSupportingText(text=texto_erro),
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="OK"),
                    on_release=lambda x: dialog.dismiss()
                )
            )
        )
        dialog.open()
        return
    
    tela_pdf = self.manager.get_screen('pdf')
    tela_pdf.descricao_cidade = campos["descricao_cidade"]
    tela_pdf.regiao_imovel = campos["regiao_imovel"]
    tela_pdf.declividade = campos["declividade"]
    tela_pdf.hidrografia = campos["hidrografia"]
    tela_pdf.resumo_solo = campos["resumo_solo"]
    tela_pdf.texto_solos = campos["texto_solos"]
    tela_pdf.data_vistoria = campos["data_vistoria"]
    tela_pdf.rotas = campos["rotas"]

    if hasattr(self, "caminho_declividade"):
        tela_pdf.caminho_declividade = self.caminho_declividade
        print(f"Caminho Declividade: {self.caminho_declividade}")
    if hasattr(self, "caminho_hidrografia"):
        tela_pdf.caminho_hidrografia = self.caminho_hidrografia
        print(f"Caminho Hidrografia: {self.caminho_hidrografia}")
    if hasattr(self, "caminho_rotas"):
        tela_pdf.caminho_rotas = self.caminho_rotas
        print(f"Caminho Rotas: {self.caminho_rotas}")
    if hasattr(self, "caminho_solos"):
        tela_pdf.caminho_solos = self.caminho_solos
        print(f"Caminho Solos: {self.caminho_solos}")

    self.manager.current_screen.manager.current = "pdf"

def preencher_cidade(self, nome_cidade):
    descricao = buscar_descricao_cidade(nome_cidade)
    self.descricao_cidade.text = descricao
    print(f"Descrição da cidade '{nome_cidade}': {descricao}")

def open_file_manager(self):
    self.file_manager = MDFileManager(
            exit_manager=lambda *args: exit_file_manager(self, *args),
            select_path=lambda path: select_path(self, path),
            preview=True,
        )
    self.file_manager.show(self.current_path)

def open_file_hidrografia(self):
    self.file_manager = MDFileManager(
            exit_manager=lambda *args: exit_file_manager(self, *args),
            select_path=lambda path: select_path_hidrografia(self, path),
            preview=True,
        )
    self.file_manager.show(self.current_path)

def open_file_rotas(self):
    self.file_manager = MDFileManager(
            exit_manager=lambda *args: exit_file_manager(self, *args),
            select_path=lambda path: select_path_rotas(self, path),
            preview=True,
        )
    self.file_manager.show(self.current_path)

def open_file_solos(self):
    self.file_manager = MDFileManager(
            exit_manager=lambda *args: exit_file_manager(self, *args),
            select_path=lambda path: select_path_solos(self, path),
            preview=True,
        )
    self.file_manager.show(self.current_path)

def exit_file_manager(self, *args):
    if hasattr(self, "file_manager") and self.file_manager:
        self.file_manager.close()

def select_path(self, path):
    exit_file_manager(self)
    self.caminho_declividade = path
    print(f"Caminho Declividade selecionado: {path}")

def select_path_hidrografia(self, path):
    exit_file_manager(self)
    self.caminho_hidrografia = path
    print(f"Caminho Hidrografia selecionado: {path}")

def select_path_rotas(self, path):
    exit_file_manager(self)
    self.caminho_rotas = path
    print(f"Caminho Rotas selecionado: {path}")

def select_path_solos(self, path):
    exit_file_manager(self)
    self.caminho_solos = path
    print(f"Caminho Solos selecionado: {path}")

def abrir_dropdown(self, *args):
    self.textos_completos = extrair_textos(self)
    self.checkboxes = {}  

    lista = MDBoxLayout(orientation="vertical", spacing=5, padding=10, size_hint_y=None)
    lista.bind(minimum_height=lista.setter("height"))
    for trecho_vermelho in self.textos_completos.keys():
        checkbox = MDListItemTrailingCheckbox(active=False)
        self.checkboxes[trecho_vermelho] = checkbox

        item = MDListItem(
            MDListItemHeadlineText(text=trecho_vermelho),
            checkbox,
            size_hint_y=None,
            height=dp(48)
        )
        lista.add_widget(item)
    
    scroll = MDScrollView(size_hint=(1,None), height=dp(300), scroll_type=['bars', 'content'])
    scroll.add_widget(lista)

    salvar_btn = MDButton(
        MDButtonText(text="Salvar"),
        on_release=lambda x: salvar_selecionados(self)
    )
    fechar_btn = MDButton(
        MDButtonText(text="Fechar"),
        on_release=lambda x: self.dialog.dismiss()
    )

    self.dialog = MDDialog(
        MDDialogHeadlineText(text="Selecione os textos desejados"),
        MDDialogContentContainer(scroll),
        MDDialogButtonContainer(salvar_btn, fechar_btn),
    )
    self.dialog.open()

def salvar_selecionados(self):
    selecionados = []
    for texto_vermelho, checkbox in self.checkboxes.items():
        if checkbox.active:
            selecionados.append(texto_vermelho)

    if selecionados:
        self.texto_solos.text = "\n\n".join(selecionados)
    else:
        self.texto_solos.text = "Nenhum texto selecionado."

    self.dialog.dismiss()

def extrair_textos(self):
        caminho = resource_path(os.path.join("models", "DECLIVIDADE e PEDOLOGIA.docx"))
        return extrair_textos_e_tabelas(caminho)
