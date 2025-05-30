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
import docx
from kivy.metrics import dp
from test2 import criar_ficha_completa

class Test(MDScreen):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        layout = MDBoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        self.text_field = MDTextField(
            MDTextFieldHintText(text="Digite algo", text_color_normal=(0, 1, 1, 1)),
            MDTextFieldHelperText(text="Este é um campo de texto"), 
            theme_text_color="Custom",
            text_color_normal=(0, 1, 1, 1),
            text_color_focus=(1, 0, 0, 1),
        )

        button = MDButton(
            MDButtonText(text="Clique aqui"),
            on_release=lambda x: criar_ficha_completa(self.text_field.text)
        )
        layout.add_widget(self.text_field)
        layout.add_widget(button)
        self.add_widget(layout)
        
class MeuApp(MDApp):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(Test(name="teste"))
        return sm

if __name__=='__main__':
    MeuApp().run()

            