from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.label import MDLabel
import os
from kivy.uix.widget import Widget
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.list import MDList, MDListItem, MDListItemHeadlineText
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.metrics import dp
from kivy.core.window import Window
from kivymd.uix.button import MDIconButton, MDButton, MDButtonText, MDButtonIcon
from app.screen2_kml.function_leitor import (
    go_back,
    go_next,
    load_directory,
)
class LeitorScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.root_path = r"H:\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
        self.current_path = self.root_path

        self.layout = MDBoxLayout(orientation="vertical")
        buttons = MDFloatLayout(size_hint_y=0.2, pos_hint={"top": 1}, height=dp(56))
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
            on_release=lambda x: go_next(self),
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Explorador de Arquivos KML",
            halign="center",
            size_hint_y=None,
            height=dp(56),
        )

        self.button = MDIconButton(
            icon="folder",
            pos_hint={"center_x": 0.5},
            on_release=lambda x: load_directory(self)
        )

        self.scroll = MDScrollView()
        self.file_list = MDList()
        self.scroll.add_widget(self.file_list)

        self.layout.add_widget(buttons) 
        self.layout.add_widget(self.label)
        self.layout.add_widget(self.button)
        self.layout.add_widget(self.scroll)

        self.add_widget(self.layout)
