import os
from kivymd.app import MDApp
from kivy.metrics import dp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDIconButton
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivymd.uix.list import MDList, MDListItem, MDListItemHeadlineText
from kivymd.uix.scrollview import MDScrollView

class LeitorScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.root_path = r"C:\Users\DESKTOP\Desktop\automacao_laudo\anexos"
        self.current_path = self.root_path

        self.layout = MDBoxLayout(orientation="vertical")

        self.label = MDLabel(
            text="Explorador de Arquivos KML",
            halign="center",
            size_hint_y=None,
            height=dp(56),
        )

        self.button = MDIconButton(
            icon="folder",
            pos_hint={"center_x": 0.5},
            on_release=self.load_directory
        )

        self.scroll = MDScrollView()
        self.file_list = MDList()
        self.scroll.add_widget(self.file_list)

        self.layout.add_widget(self.label)
        self.layout.add_widget(self.button)
        self.layout.add_widget(self.scroll)

        self.add_widget(self.layout)

        self.load_directory()

    def load_directory(self, *args):
        self.file_list.clear_widgets()

        # Adiciona botão de voltar se não estiver no diretório raiz
        if self.current_path != self.root_path:
            voltar_item = MDListItem(
                on_release=self.go_up
            )
            voltar_item.add_widget(MDListItemHeadlineText(text=".. (voltar)"))
            self.file_list.add_widget(voltar_item)

        try:
            itens = os.listdir(self.current_path)
        except FileNotFoundError:
            itens = []

        for nome in sorted(itens):
            caminho = os.path.join(self.current_path, nome)

            if os.path.isdir(caminho):
                item = MDListItem(
                    on_release=lambda x, p=caminho: self.entrar_em_pasta(p)
                )
                item.add_widget(MDListItemHeadlineText(text=f"[DIR] {nome}"))
                self.file_list.add_widget(item)

            elif nome.lower().endswith(".kml"):
                item = MDListItem(
                    on_release=lambda x, f=caminho: self.on_file_selected(f)
                )
                item.add_widget(MDListItemHeadlineText(text=nome))
                self.file_list.add_widget(item)

    def entrar_em_pasta(self, pasta):
        self.current_path = pasta
        self.load_directory()

    def go_up(self, *args):
        self.current_path = os.path.dirname(self.current_path)
        self.load_directory()

    def on_file_selected(self, caminho):
        MDSnackbar(
            MDSnackbarText(text=f"Selecionado:\n{caminho}"),
            y=dp(24),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.8,
        ).open()
        # Aqui você pode processar o .kml selecionado


class MainApp(MDApp):
    def build(self):
        self.title = "Leitor KML"
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Dark"
        return LeitorScreen()

if __name__ == "__main__":
    MainApp().run()
