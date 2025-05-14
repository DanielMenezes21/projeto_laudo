from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.button import MDIconButton, MDButton, MDButtonIcon, MDButtonText
from kivymd.uix.progressindicator import MDCircularProgressIndicator
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivy.clock import Clock
from kivy.core.window import Window
from threading import Thread
from kivy.metrics import dp
from kivy.utils import get_color_from_hex
from kivymd.uix.dialog import MDDialog
from app.screen1_email.hp_function import (
    next_screen,
    download_file,
    after_download,
    show_file_manager,

)

class MenuScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        
        self.file_manager = None

        layout = MDFloatLayout(size_hint=(1, 1))

        self.text_field = MDTextField(
            MDTextFieldTrailingIcon(icon="magnify"),
            MDTextFieldHintText(
                text="palavra chave", 
                theme_text_color="Custom", 
                text_color_normal="yellow", 
                text_color_focus="yellow"
            ),
            MDTextFieldHelperText(
                text="a automação irá buscar os emails com essa palavra chave", 
                mode="on_focus", 
                theme_text_color="Custom", 
                text_color_normal="yellow", 
                text_color_focus="yellow"
            ),
            MDTextFieldTrailingIcon(
                icon="magnify"
            ),
            mode="outlined", 
            size_hint_x=None, 
            width=dp(440), 
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            pos_hint={"center_x": 0.5, "center_y": 0.7}
        )
        
        self.download_button = MDButton(
            MDButtonIcon(icon="download", icon_color="white"),
            MDButtonText(text="Baixar arquivos", text_color="white"),
            pos_hint={"center_x": 0.1, "y": 0.5},
            size_hint=(None, None),
            size=(dp(56), dp(56)),
            on_release=lambda x: download_file(self, x),
        )

        self.file_manager_button = MDButton(
            MDButtonIcon(icon="folder", icon_color="white"),
            MDButtonText(text="Gerenciar arquivos", text_color="white"),
            pos_hint={"center_x": 0.5, "y": 0.5},
            size_hint=(None, None),
            size=(dp(56), dp(56)),
            on_release=lambda x: show_file_manager(self),
        )

        self.progress = MDCircularProgressIndicator(
            size_hint=(None, None),
            size=(dp(56), dp(56)),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            active=False,
            color="white",
            theme_text_color="Custom"
        )

        self.button_next = MDButton(
            MDButtonIcon(icon="arrow-right"),
            MDButtonText(text="Next"),
            pos_hint={"x": 0.9, "y": 0.5},
            on_release=lambda x: next_screen(self, x),
        )

        layout.add_widget(self.download_button)
        layout.add_widget(self.file_manager_button)
        layout.add_widget(self.text_field)
        layout.add_widget(self.progress)
        layout.add_widget(self.button_next)

        self.add_widget(layout)

    