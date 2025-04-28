from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.dropdownitem import MDDropDownItem, MDDropDownItemText
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from app.screen3_dadosp.dados_function import extrair_dados_pdf, go_back, go_next, abrir_seletor_pdf, fechar_arquivo, on_pdf_selecionado
from kivy.uix.widget import Widget
from kivy.metrics import dp
import os
import re
from docx import Document

class DadosScreen(MDScreen):
    def open_dropdown(self, *args):
        self.dropdown.open()

    def set_tratamento(self, valor):
        self.tratamento = valor
        self.botao.children[0].text = valor  # atualiza o texto do botão
        self.dropdown.dismiss()

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.tratamento = ""

        self.scroll = MDScrollView()

        self.layout = MDBoxLayout(orientation='vertical', padding=10, spacing=20, size_hint_y=None)
        self.layout.bind(minimum_height=self.layout.setter('height'))
        buttons = MDFloatLayout(size_hint_y=None)
        buttons.add_widget(Widget())
        cliente = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))
        coord = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_back(self),
        )
        buttons.add_widget(self.button_back)

        self.button_next = MDIconButton(
            icon="arrow-right",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_next(self),
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Dados do Proponente",
            halign="center",
            theme_text_color="Custom",
            text_color="yellow",
            size_hint=(1, None),
            pos_hint={"center_x": 0.5, "center_y":0.9}
        )

        self.agencia = MDTextField(
            MDTextFieldHintText(text="Agencia"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.botao = MDButton(
            MDButtonText(text="Selecionar Tratamento"),
            pos_hint={"center_x": 0.5},
            on_release=self.open_dropdown
        )
        cliente.add_widget(self.botao)
        
        self.proponente = MDTextField(
            MDTextFieldHintText(text="Proponente"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )
        cliente.add_widget(self.proponente)

        self.matricula = MDTextField(
            MDTextFieldHintText(text="Matricula"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.cpf = MDTextField(
            MDTextFieldHintText(text="CPF"),
            size = (200,50),
            size_hint=(0.9, None),
            write_tab=False,
            pos_hint={"center_x": 0.5},
            height=50,
        )

        self.nome_imovel = MDTextField(
            MDTextFieldHintText(text="Nome do Imóvel"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.civil = MDTextField(
            MDTextFieldHintText(text="situação civil"),
            MDTextFieldHelperText(text="fale sobre a situação civil do proponente, se o mesmo se encontra casado,\n solteiro, viuvo ou se outra pessoa partilha a terra com o mesmo"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.municipio = MDTextField(
            MDTextFieldHintText(text="Município do imóvel"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.estado = MDTextField(
            MDTextFieldHintText(text="Estado do imóvel"),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.latitude=MDTextField(
            MDTextFieldHintText(text="Latitude do centróide do imóvel"),
            size=(100,50),
            size_hint=(0.4, None),
            write_tab=False,
            pos_hint={"x": 0.05}
        )
        coord.add_widget(self.latitude)

        self.longitude=MDTextField(
            MDTextFieldHintText(text="Longitude do centróide do imóvel"),
            size=(100,50),
            size_hint=(0.4, None),
            write_tab=False,
            pos_hint={"x":0.6}
        )
        coord.add_widget(self.longitude)

        self.botao_selecionar = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: abrir_seletor_pdf(self),
        )
        self.botao_selecionar.add_widget(MDButtonText(text="Selecionar PDF"))

        self.file_manager = MDFileManager(
            exit_manager=lambda x: fechar_arquivo(self),
            select_path=lambda x: on_pdf_selecionado(self, x),
            ext=[".pdf"]
        )

        self.dropdown = MDDropdownMenu(
            caller=self.botao,
            items=[
                {"text": "Sr.", "on_release": lambda x="Sr.": self.set_tratamento(x)},
                {"text": "Srª", "on_release": lambda x="Srª": self.set_tratamento(x)},
            ],
            width_mult=3
        )  

        self.layout.add_widget(buttons)
        self.layout.add_widget(self.label)
        self.layout.add_widget(self.agencia)
        self.layout.add_widget(self.matricula)
        self.layout.add_widget(self.civil)
        self.layout.add_widget(cliente)
        self.layout.add_widget(self.cpf)
        self.layout.add_widget(self.nome_imovel)
        self.layout.add_widget(self.municipio)
        self.layout.add_widget(self.estado)
        self.layout.add_widget(coord)
        self.layout.add_widget(self.botao_selecionar)
        
        self.scroll.add_widget(self.layout)

        self.add_widget(self.scroll)
