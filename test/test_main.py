from kivymd.app import MDApp
from kivymd.uix.screenmanager import ScreenManager
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField, MDTextFieldHelperText, MDTextFieldHintText
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.list import MDList, MDListItem, MDListItemHeadlineText, MDListItemTrailingCheckbox
from kivymd.uix.expansionpanel import MDExpansionPanel, MDExpansionPanelContent, MDExpansionPanelHeader
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.dialog import MDDialog, MDDialogButtonContainer, MDDialogContentContainer, MDDialogHeadlineText
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from kivy.metrics import dp
from test_create import gerar_documento_completo
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.utils import platform

class Test(MDScreen):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        layout = MDBoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))

        self.text_field = MDTextField(
            MDTextFieldHintText(text="Valor"),
            theme_text_color="Custom",
            text_color_normal=(0, 1, 1, 1),
            text_color_focus=(1, 0, 0, 1),
        )

        self.quant = MDTextField(
            MDTextFieldHintText(text="simulação de quantidade de matrículas"),
            theme_text_color="Custom",
            text_color_normal=(0, 1, 1, 1),
            text_color_focus=(1, 0, 0, 1),
        )

        checkbox_layout = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40), spacing=dp(10))
        self.checkbox = MDCheckbox()
        self.checkbox.bind(active=self.on_checkbox_active)
        self.celula_verde = False
        checkbox_label = MDListItemHeadlineText(text="Opção de exemplo")
        checkbox_layout.add_widget(self.checkbox)
        checkbox_layout.add_widget(checkbox_label)

        self.selected_image = None
        image_button = MDButton(
            MDButtonText(text="Selecionar Imagem"),
            on_release=self.open_file_manager
        )

        button = MDButton(
            MDButtonText(text="Clique aqui"),
            on_release=lambda x: gerar_documento_completo(
                self.text_field.text,
                f"LAUDO DE AVALIAÇÃO Nº {self.text_field.text},\n 01 de Janeiro de 2024, PALMAS TO",
                celula_verde=self.checkbox.active,
                imagem_path=self.selected_image,
                quantidade=int(self.quant.text) if self.quant.text.isdigit() else 1
            )     
        )

        layout.add_widget(self.text_field)
        layout.add_widget(self.quant)
        layout.add_widget(checkbox_layout)
        layout.add_widget(image_button)
        layout.add_widget(button)
        self.add_widget(layout)

        self.file_manager = MDFileManager(
            select_path=self.select_path,
            exit_manager=self.exit_manager,
            preview=True
        )

    def open_file_manager(self, *args):
        start_path = "/" if platform == "linux" else "C:\\Users\\DESKTOP\\Desktop\\automacao_laudo"
        self.file_manager.show(start_path)

    def select_path(self, path):
        self.selected_image = path
        self.exit_manager()

    def exit_manager(self, *args):
        self.file_manager.close()

    def on_checkbox_active(self, checkbox, value):
        self.celula_verde = value  

class MeuApp(MDApp):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(Test(name="teste"))
        return sm

if __name__=='__main__':
    MeuApp().run()

            