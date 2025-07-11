from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDIconButton, MDButton, MDButtonText, MDButtonIcon
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.list import MDList
from kivymd.uix.boxlayout import MDBoxLayout
from kivy.metrics import dp
import os
from datetime import datetime
from kivy.uix.widget import Widget
from kivymd.app import MDApp
from modules.data_atual import formatar_data

from app.screen5_insertpdf.pdf_function import (
    go_back,
    initialize_word,
    abrir_gerenciador,
)
from app.screen5_insertpdf.pdf_insercao import gerar_documento


class PDFInsert(MDScreen):
    def receber_dados_matriculas(self, lista_dados):
        self.lista_dados_matriculas = lista_dados
        self.qtd_imoveis = len(lista_dados)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.lista_dados_matriculas = []  
        self.qtd_imoveis = 0
        self.pastas_abertas = set()
        data = formatar_data()
        data_nome = datetime.now()
        mes = f'{data_nome.month:02d}. {data.split('de')[1].strip()}'
        self.root_path = r"\\10.0.100.160\\Agropassos\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
        self.root_path = os.path.join(self.root_path, mes)
        if not os.path.exists(self.root_path):
            self.root_path = os.path.join(os.path.expanduser("~/Documents"), mes)
        self.current_path = self.root_path

        self.layout = MDBoxLayout(orientation="vertical")
        buttons = MDFloatLayout(size_hint_y=None)
        buttons.add_widget(Widget())

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
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Explorador de Arquivos PDF",
            halign="center",
            size_hint_y=None,
            height=dp(56),
        )

        self.button = MDButton(
            MDButtonIcon(
                icon="folder",
                theme_icon_color="Custom",
                icon_color=(1, 1, 1, 1),
            ),
            MDButtonText(
                text="Selecionar PDF do CAR",
                theme_text_color="Custom",
                text_color=(1, 1, 1, 1),
            ),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            size_hint=(0.5, None),
            on_release=lambda x: abrir_gerenciador(self, tipo="car")
        )

        self.button2 = MDButton(
            MDButtonIcon(
                icon="folder",
                theme_icon_color="Custom",
                icon_color=(1, 1, 1, 1),
            ),
            MDButtonText(
                text="Selecionar PDF do CIT",
                theme_text_color="Custom",
                text_color=(1, 1, 1, 1),
            ),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            size_hint=(0.5, None),
            on_release=lambda x: abrir_gerenciador(self, tipo="cit")
        )

        self.button_gerar = MDButton(
            MDButtonIcon(
                icon="file-word",
                theme_icon_color="Custom",
                icon_color=(1, 1, 1, 1),
            ),
            MDButtonText(
                text="Gerar Documento Word",
                theme_text_color="Custom",
                text_color=(1, 1, 1, 1),
            ),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            size_hint=(0.5, None),
            on_release=lambda x: gerar_documento(self)
        )
        

        self.scroll = MDScrollView()
        self.file_list = MDList()
        self.scroll.add_widget(self.file_list)

        self.scroll2 = MDScrollView()
        self.file_list2 = MDList()
        self.scroll2.add_widget(self.file_list2)

        self.layout.add_widget(buttons)
        self.layout.add_widget(self.label)
        self.layout.add_widget(self.button)
        self.layout.add_widget(self.button2)
        self.layout.add_widget(self.scroll)
        self.layout.add_widget(self.scroll2)
        self.layout.add_widget(self.button_gerar)

        self.add_widget(self.layout)
        self.caminho_car = ""
        self.caminho_cit = ""
        initialize_word(self)

   