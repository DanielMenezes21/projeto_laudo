from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.dialog import MDDialog, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.button import MDButton, MDButtonText
from kivy.uix.image import Image
from kivymd.uix.menu import MDDropdownMenu
from docx import Document
from modules.pesquisa import buscar_descricao_cidade
from modules.resource_path import resource_path
import os

def go_back(self):
    self.manager.current_screen.manager.current = "dados"

def go_next1(self):
    campos = {
        "descricao_imovel": self.descricao_imovel.text,
        "descricao_cidade": self.descricao_cidade.text,
        "atividade_imovel": self.atividade_imovel.text,
        "regiao_imovel": self.regiao_imovel.text,
        "declividade": self.declividade_text.text,
        "hidrografia": self.hidrografia_text.text,
        "resumo_solo": self.resumo_solo.text,
        "texto_solos": self.texto_solos.text,
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
    tela_pdf.descricao_imovel = campos["descricao_imovel"]
    tela_pdf.descricao_cidade = campos["descricao_cidade"]
    tela_pdf.atividade_imovel = campos["atividade_imovel"]
    tela_pdf.regiao_imovel = campos["regiao_imovel"]
    tela_pdf.declividade = campos["declividade"]
    tela_pdf.hidrografia = campos["hidrografia"]
    tela_pdf.resumo_solo = campos["resumo_solo"]
    tela_pdf.texto_solos = campos["texto_solos"]
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

def abrir_dropdown(self, *args):
    self.textos_completos = extrair_textos(self)
    menu_items = [
        {
            "text": trecho_vermelho,
            "on_release": lambda x=trecho_vermelho: selecionar_opcao(self, x),
        }
        for trecho_vermelho in self.textos_completos.keys()
    ]

    self.menu = MDDropdownMenu(
        caller=self.opcao_solo,
        items=menu_items,
        width_mult=4,
    )
    self.menu.open()

def selecionar_opcao(self, texto_vermelho):
    texto_completo = self.textos_completos.get(texto_vermelho, texto_vermelho)
    self.texto_solos.text = texto_completo
    if self.menu:
        self.menu.dismiss()

def extrair_textos(self):
        caminho = resource_path(os.path.join("models", "DECLIVIDADE e PEDOLOGIA.docx"))
        doc = Document(caminho)

        resultados = {}
        for par in doc.paragraphs:
            runs = par.runs
            i = 0
            while i < len(runs):
                run = runs[i]
                if run.font.color and run.font.color.rgb and str(run.font.color.rgb) == "FF0000":
                    texto_vermelho = run.text.strip()
                    texto_completo = texto_vermelho
                    i += 1
                    while i < len(runs):
                        next_text = runs[i].text
                        if "#" in next_text:
                            texto_completo += " " + next_text.split("#")[0]
                            break
                        texto_completo += " " + next_text
                        i += 1
                    if texto_vermelho:
                        resultados[texto_vermelho] = texto_completo.strip()
                else:
                    i += 1

        if not resultados:
            resultados["Nenhum texto em vermelho encontrado."] = ""

        return resultados
