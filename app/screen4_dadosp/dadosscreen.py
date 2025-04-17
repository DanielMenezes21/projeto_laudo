from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from app.screen4_dadosp.dados_function import extrair_dados_pdf
import os

class DadosScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.layout = MDFloatLayout()
        self.add_widget(self.layout)

        self.label = MDLabel(
            text="Dados do Proponente",
            halign="center",
            theme_text_color="Custom",
            text_color="yellow",
            size_hint=(1, 0.1),
            pos_hint={"center_x": 0.5, "y": 0.9},
        )
        self.layout.add_widget(self.label)

        self.proponente = MDTextField(
            MDTextFieldHintText(text="Proponente"),
            pos_hint={"center_x": 0.5, "y": 0.75},
            size_hint=(0.8, None),
            height=50,
        )
        self.layout.add_widget(self.proponente)

        self.cpf = MDTextField(
            MDTextFieldHintText(text="CPF"),
            pos_hint={"center_x": 0.5, "y": 0.65},
            size_hint=(0.8, None),
            height=50,
        )
        self.layout.add_widget(self.cpf)

        self.botao_selecionar = MDButton(
            pos_hint={"center_x": 0.5, "y": 0.5},
            on_release=self.abrir_seletor_pdf,
        )
        self.botao_selecionar.add_widget(MDButtonText(text="Selecionar PDF"))
        self.layout.add_widget(self.botao_selecionar)

        self.file_manager = MDFileManager(
            exit_manager=self.fechar_arquivo,
            select_path=self.carregar_dados_pdf,
            ext=[".pdf"]
        )

    def abrir_seletor_pdf(self, *args):
        self.file_manager.show(os.getcwd())

    def fechar_arquivo(self, *args):
        self.file_manager.close()

    def carregar_dados_pdf(self, caminho_pdf):
        self.fechar_arquivo()
        nome, cpf = extrair_dados_pdf(caminho_pdf)
        self.proponente.text = nome
        self.cpf.text = cpf
