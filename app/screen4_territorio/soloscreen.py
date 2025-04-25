from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import ScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.uix.widget import Widget
from kivy.metrics import dp

class SoloScreen(MDScreen):
    def __init__(self, **Kwargs):
        super.__init__(**Kwargs)

        self.scroll = ScrollView(1,1)

        self.layout = MDBoxLayout(orientation="vertical", size_hint_y=None, padding = 20)
        self.layout.bind(minimum_height=self.layout.setter("height"))
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
            on_release=lambda x: go_next(self)
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Caracteristicas do solo e cidade",
            theme_text_color = "Custom",
            text_color = "yellow",
            size_hint=(1,None),
            pos_hint={"center_x":0.5, "center_y":0.9}
        )

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
