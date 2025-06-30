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
from app.screen3_dadosp.dados_function import receber_dados_pdf, go_back, go_next, abrir_seletor_pdf, fechar_arquivo, on_pdf_selecionado
from kivy.uix.widget import Widget
from kivy.metrics import dp
from kivy.clock import Clock
from app.screen3_dadosp.dados_function import abrir_dialogo_matriculas
import os
import re
from docx import Document

class DadosScreen(MDScreen):
    def open_dropdown(self, *args):
        self.dropdown.open()

    def set_tratamento(self, valor):
        self.tratamento = valor
        self.botao.children[0].text = valor
        self.dropdown.dismiss()
        if self.proponente_atual in self.proponentes:
            self.proponentes[self.proponente_atual]["tratamento"] = valor
    
    def salvar_civil(self, instance, value):
        if not value and self.proponente_atual in self.proponentes:
            self.proponentes[self.proponente_atual]["civil"] = instance.text

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.qtd_imoveis = 0
        self.tratamento = ""

        self.scroll = MDScrollView(bar_color=(1, 1, 1, 0.5), bar_width=10, scroll_type=["bars", "content"])

        self.layout = MDBoxLayout(orientation='vertical', padding=15, spacing=20, size_hint_y=None)
        self.layout.bind(minimum_height=self.layout.setter('height'))
        buttons = MDFloatLayout(size_hint_y=None)
        buttons.add_widget(Widget())
        cliente = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))

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

        self.solicitante = MDTextField(
            MDTextFieldHintText(text="Solicitante"),
            MDTextFieldHelperText(text="O solicitante que enviou os documentos para o laudo",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
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
        
        self.proponentes = []
        self.proponente_atual = ""
        self.proponente = MDTextField(
            MDTextFieldHintText(text="Proprietario"),
            MDTextFieldHelperText(text="Nome do proprietário",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )
        cliente.add_widget(self.proponente)

        self.botao_matricula = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: abrir_dialogo_matriculas(self)
        )
        self.botao_matricula.add_widget(MDButtonText(text="Definir Matrículas"))

        self.cpf = MDTextField(
            MDTextFieldHintText(text="CPF",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            write_tab=False,
            pos_hint={"center_x": 0.5},
            height=50,
        )

        self.civil = MDTextField(
            MDTextFieldHintText(text="situação civil"),
            MDTextFieldHelperText(text="fale sobre a situação civil do proponente, se o mesmo se encontra casado,\n solteiro, viuvo ou se outra pessoa partilha a terra com o mesmo",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.municipio = MDTextField(
            MDTextFieldHintText(text="Município do imóvel"),
            MDTextFieldHelperText(text="Município onde o imóvel está localizado",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.estado = MDTextField(
            MDTextFieldHintText(text="Estado do imóvel"),
            MDTextFieldHelperText(text="Estado onde o imóvel está localizado",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.botao_proponente = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: self.menu_proponente.open()
        )
        self.botao_proponente.add_widget(MDButtonText(text="Selecionar Proponente"))

        self.menu_proponente = MDDropdownMenu(
            caller=self.botao_proponente,
            items=[],  
            width_mult=4,
        )
        cliente.add_widget(self.botao_proponente)

        self.botao_selecionar = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: abrir_seletor_pdf(self),
        )
        self.botao_selecionar.add_widget(MDButtonText(text="Selecionar PDF"))

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
        self.layout.add_widget(self.solicitante)
        self.layout.add_widget(self.botao_matricula)
        self.layout.add_widget(self.civil)
        self.layout.add_widget(cliente)
        self.layout.add_widget(self.cpf)
        self.layout.add_widget(self.municipio)
        self.layout.add_widget(self.estado)
        self.layout.add_widget(self.botao_selecionar)

        self.solicitante.bind(focus=self._on_focus)
        self.civil.bind(focus=self._on_focus)
        self.cpf.bind(focus=self._on_focus)
        self.municipio.bind(focus=self._on_focus)
        self.estado.bind(focus=self._on_focus)
        self.civil.bind(focus=self.salvar_civil)
        
        self.scroll.add_widget(self.layout)

        self.add_widget(self.scroll)

        Window.bind(on_keyboard_height=self._ajustar_scroll)
        self.campo_em_foco = None

    def _on_focus(self, instance, value):
        if value: 
            self.campo_em_foco = instance
            Clock.schedule_once(lambda dt: self.scroll.scroll_to(instance), 0.1)

    def _ajustar_scroll(self, window, altura_teclado):
        if altura_teclado > 0 and self.campo_em_foco:
            Clock.schedule_once(lambda dt: self.scroll.scroll_to(self.campo_em_foco), 0.1)
        else:
            self.scroll.scroll_y = 1

    def selecionar_proprietario(self, nome_escolhido):
        self.menu_proponente.dismiss()
        self.proponente_atual = nome_escolhido

        dados = self.proponentes[nome_escolhido]
        self.proponente.text = nome_escolhido
        self.cpf.text = dados.get("cpf", "")
        self.civil.text = dados.get("civil", "")
        self.set_tratamento(dados.get("tratamento", ""))

