from kivymd.uix.screen import MDScreen
from kivymd.app import MDApp
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.clock import Clock
import os
from kivy.uix.widget import Widget
from kivy.metrics import dp
from kivy.core.window import Window
from app.screen4_territorio.solofunction import (
    go_back, go_next1, open_file_manager, 
    open_file_hidrografia, preencher_cidade, 
    abrir_dropdown, open_file_rotas, open_file_solos)

class SoloScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.root = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL"
        self.current_path = self.root
        if not os.path.exists(self.root):
            self.root = os.path.join(os.path.expanduser("~/Documents"))
        Window.bind(on_key_down=self._verifica_enter)

        self.scroll = MDScrollView(bar_color=(1, 1, 1, 0.5), bar_width=10, scroll_type=["bars", "content"])

        self.layout = MDBoxLayout(orientation="vertical", size_hint_y=None, padding = 20, spacing = 20)
        self.layout.bind(minimum_height=self.layout.setter("height"))
        buttons = MDFloatLayout(size_hint_y=None)
        buttons.add_widget(Widget())
        declividade = MDBoxLayout(orientation="horizontal", size_hint_y=None, padding = 20)
        declividade.bind(minimum_height=declividade.setter("height"))
        hidrografia = MDBoxLayout(orientation="horizontal", size_hint_y=None, padding = 20)
        hidrografia.bind(minimum_height=hidrografia.setter("height"))
        solo = MDBoxLayout(orientation="horizontal", size_hint_y=None, padding = 20)
        solo.bind(minimum_height=solo.setter("height"))
        rotas = MDBoxLayout(orientation="horizontal", size_hint_y=None, padding = 20)
        rotas.bind(minimum_height=rotas.setter("height"))

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.8},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_back(self),
        )
        buttons.add_widget(self.button_back)

        self.button_next = MDIconButton(
            icon="arrow-right",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.8},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_next1(self)
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Caracteristicas do solo e cidade",
            theme_text_color = "Custom",
            text_color = "yellow",
            size_hint=(1,None),
            pos_hint={"center_x":0.5, "center_y":0.9}
        )

        self.descricao_imovel = MDTextField(
            MDTextFieldHintText(text="Descrição do imovel"),
            MDTextFieldHelperText(text="Ex: corregos, grotas, represas, cachoeiras, etc."),
            size_hint=(1, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.8},
            multiline=False,
        )

        self.descricao_cidade = MDTextField(
            MDTextFieldHintText(text="Descrição da cidade"),
            MDTextFieldHelperText(text="aperte a tecla 'Enter' para preencher"),
            MDTextFieldTrailingIcon(icon="magnify"),
            size_hint=(1, None),
            write_tab=False,
            height=dp(150),
            pos_hint={"center_x": 0.5, "center_y": 0.7},
            multiline=True,
            on_text_validate=lambda x: preencher_cidade(self, self.descricao_cidade.text)
        )

        self.regiao_imovel = MDTextField(
            MDTextFieldHintText(text="Região do imovel"),
            MDTextFieldHelperText(text="Ex: reconhecimento em plantio, relevo."),
            size_hint=(1, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.7},
            multiline=False,
        )

        self.atividade_imovel = MDTextField(
            MDTextFieldHintText(text="Atividade do imovel"),
            MDTextFieldHelperText(text="Ex: Pecuaria, agricultura, etc."),
            size_hint=(1, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.6},
            multiline=False,
        )

        self.declividade_text = MDTextField(
            MDTextFieldHintText(text="Declividade do imovel"),
            MDTextFieldHelperText(text="Descrição da declividade do terreno"),
            size_hint=(0.8, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            multiline=False,
        ) 
        declividade.add_widget(self.declividade_text)

        self.imagem_declividade = MDIconButton(
            icon="image",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.5},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release = lambda x: open_file_manager(self)
        )
        declividade.add_widget(self.imagem_declividade)

        self.hidrografia_text = MDTextField(
            MDTextFieldHintText(text="Hidrografia do imovel"),
            MDTextFieldHelperText(text="Descrição da hidrografia do terreno"),
            size_hint=(0.8, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            size_hint_y=None,
            multiline=False,
        )
        hidrografia.add_widget(self.hidrografia_text)

        self.imagem_hidrografia = MDIconButton(
            icon="image",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.5},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release = lambda x: open_file_hidrografia(self)
        )
        hidrografia.add_widget(self.imagem_hidrografia)

        self.opcao_solo = MDButton(
            MDButtonText(text="Tipos de solo"),
            size_hint=(0.2, None),
            pos_hint={"center_x": 0.5, "center_y": 0.4},
            on_release=lambda x: abrir_dropdown(self),
        )
        solo.add_widget(self.opcao_solo)

        self.resumo_solo = MDTextField(
            MDTextFieldHintText(text="Resumo do solo"),
            MDTextFieldHelperText(text="identificação do solo da propriedade"),
            size_hint=(0.8, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.4},
            multiline=False,
        )
        solo.add_widget(self.resumo_solo)

        self.imagem_solos = MDIconButton(
            icon="image",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: open_file_solos(self)
        )
        solo.add_widget(self.imagem_solos)

        self.texto_solos = MDTextField(
            MDTextFieldHintText(text="Texto completo selecionado"),
            MDTextFieldHelperText(text="Texto que foi selecionado no botão 'Tipo de solo'"),
            size_hint=(1, None),
            write_tab=False,
            height=dp(150),
            pos_hint={"center_x": 0.5, "center_y": 0.4},
            multiline=True,
        )

        self.rotas_text = MDTextField(
            MDTextFieldHintText(text="Rotas do imovel"),
            MDTextFieldHelperText(text="Descrição da rota de acesso ao imovel"),
            size_hint=(1, None),
            write_tab=False,
            height=dp(40),
            pos_hint={"center_x": 0.5, "center_y": 0.4},
            multiline=False,
        )
        rotas.add_widget(self.rotas_text)

        self.rotas_imagem = MDIconButton(
            icon="image",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release = lambda x: open_file_rotas(self)
        )
        rotas.add_widget(self.rotas_imagem)
        
        self.caminho_declividade = ''
        self.caminho_hidrografia = ''
        self.caminho_rotas = ''
        self.caminho_solos = ''
        self.menu = None
        self.textos_completos = {}
        
        self.layout.add_widget(buttons)
        self.layout.add_widget(self.label)
        self.layout.add_widget(self.descricao_imovel)
        self.layout.add_widget(self.descricao_cidade)
        self.layout.add_widget(self.atividade_imovel)
        self.layout.add_widget(self.regiao_imovel)
        self.layout.add_widget(declividade)
        self.layout.add_widget(hidrografia)
        self.layout.add_widget(solo)
        self.layout.add_widget(self.texto_solos)
        self.layout.add_widget(rotas)

        self.scroll.add_widget(self.layout)
        self.add_widget(self.scroll)
        
    def _verifica_enter(self, window, key, scancode, codepoint, modifier):
        if key == 13 and self.descricao_cidade.focus:
            from app.screen4_territorio.solofunction import buscar_descricao_cidade
            nome_cidade = self.descricao_cidade.text.strip()
            if nome_cidade:
                descricao = buscar_descricao_cidade(nome_cidade)
                self.descricao_cidade.text = descricao
            return True
        return False
    