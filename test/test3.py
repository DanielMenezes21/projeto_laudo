from kivymd.app import MDApp
from kivy.core.window import Window
from kivymd.uix.screen import MDScreen
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField
from kivy.metrics import dp
from docx import Document


class TelaTeste3(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.menu = None
        self.textos_completos = {}

        self.text_field = MDTextField(
            hint_text="Texto completo selecionado",
            multiline=True,
            size_hint=(0.8, None),
            height=dp(150),
            pos_hint={"center_x": 0.5, "center_y": 0.6},
        )
        self.add_widget(self.text_field)

        self.botao = MDButton(
            MDButtonText(text="Selecionar trecho em vermelho"),
            pos_hint={"center_x": 0.5, "center_y": 0.4},
            on_release=self.abrir_dropdown,
        )
        self.add_widget(self.botao)

    def abrir_dropdown(self, *args):
        self.textos_completos = self.extrair_textos()
        menu_items = [
            {
                "text": trecho_vermelho,
                "on_release": lambda x=trecho_vermelho: self.selecionar_opcao(x),
            }
            for trecho_vermelho in self.textos_completos.keys()
        ]

        self.menu = MDDropdownMenu(
            caller=self.botao,
            items=menu_items,
            width_mult=4,
        )
        self.menu.open()

    def selecionar_opcao(self, texto_vermelho):
        texto_completo = self.textos_completos.get(texto_vermelho, texto_vermelho)
        self.text_field.text = texto_completo
        if self.menu:
            self.menu.dismiss()

    def extrair_textos(self):
        caminho = r"models\DECLIVIDADE e PEDOLOGIA.docx"
        doc = Document(caminho)

        resultados = {}
        for par in doc.paragraphs:
            runs = par.runs
            i = 0
            while i < len(runs):
                run = runs[i]
                if run.font.color and run.font.color.rgb and str(run.font.color.rgb) == "FF0000":
                    texto_vermelho = run.text.strip()
                    texto_completo = texto_vermelho
                    i += 1
                    while i < len(runs):
                        next_text = runs[i].text
                        if "#" in next_text:
                            texto_completo += " " + next_text.split("#")[0]
                            break
                        texto_completo += " " + next_text
                        i += 1
                    if texto_vermelho:
                        resultados[texto_vermelho] = texto_completo.strip()
                else:
                    i += 1

        if not resultados:
            resultados["Nenhum texto em vermelho encontrado."] = ""

        return resultados

class Test(MDApp):
    def build(self):
        Window.size = (400, 600)
        return TelaTeste3()

if __name__ == "__main__":
    Test().run()