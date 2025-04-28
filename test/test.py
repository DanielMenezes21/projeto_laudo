from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.textfield import MDTextField, MDTextFieldHelperText
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.label import MDLabel
from kivymd.uix.menu import MDDropdownMenu
from kivy.metrics import dp


class DadosScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        scroll = MDScrollView()
        self.layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20, size_hint_y=None)
        self.layout.bind(minimum_height=self.layout.setter("height"))

        # === SEÇÃO: PROONENTE ===
        self.layout.add_widget(MDLabel(text="Dados do Proponente", halign="center", bold=True))
        cliente = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))

        # Tratamento
        self.tratamento = ""
        self.tratamento_botao = MDButton(
            MDButtonText(text="Tratamento"),
            size_hint_x=0.3,
            on_release=self.abrir_dropdown
        )
        cliente.add_widget(self.tratamento_botao)

        self.dropdown = MDDropdownMenu(
            caller=self.tratamento_botao,
            items=[
                {"text": "Sr.", "on_release": lambda x="Sr.": self.set_tratamento(x)},
                {"text": "Srª", "on_release": lambda x="Srª": self.set_tratamento(x)},
            ],
            width_mult=3
        )

        # Nome e CPF
        self.proponente = MDTextField(
            hint_text="Nome completo",
            size_hint_x=0.7,
            height=dp(48)
        )
        cliente.add_widget(self.proponente)

        self.layout.add_widget(cliente)
        self.cpf = MDTextField(hint_text="CPF", size_hint_y=None, height=dp(48))
        self.layout.add_widget(self.cpf)

        # Situação civil
        self.civil = MDTextField(
            MDTextFieldHelperText(text="Casado, solteiro, etc."),
            hint_text="Situação civil",
            
            size_hint_y=None,
            height=dp(48)
        )
        self.layout.add_widget(self.civil)

        # === SEÇÃO: IMÓVEL ===
        self.layout.add_widget(MDLabel(text="Dados do Imóvel", halign="center", bold=True))

        self.nome_imovel = MDTextField(hint_text="Nome do imóvel", size_hint_y=None, height=dp(48))
        self.matricula = MDTextField(hint_text="Matrícula", size_hint_y=None, height=dp(48))
        self.municipio = MDTextField(hint_text="Município", size_hint_y=None, height=dp(48))
        self.estado = MDTextField(hint_text="Estado", size_hint_y=None, height=dp(48))

        self.layout.add_widget(self.nome_imovel)
        self.layout.add_widget(self.matricula)
        self.layout.add_widget(self.municipio)
        self.layout.add_widget(self.estado)

        # === SEÇÃO: COORDENADAS ===
        self.layout.add_widget(MDLabel(text="Coordenadas", halign="center", bold=True))

        linha_coords = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))
        self.latitude = MDTextField(hint_text="Latitude", size_hint_y=None, height=dp(48))
        self.longitude = MDTextField(hint_text="Longitude", size_hint_y=None, height=dp(48))
        linha_coords.add_widget(self.latitude)
        linha_coords.add_widget(self.longitude)
        self.layout.add_widget(linha_coords)

        # === BOTÃO DE CONTINUAÇÃO ===
        self.botao_continuar = MDButton(
            MDButtonText(text="Continuar"),
            pos_hint={"center_x": 0.5},
            size_hint=(None, None),
            size=(dp(150), dp(48)),
            on_release=self.proxima_tela
        )
        self.layout.add_widget(self.botao_continuar)

        scroll.add_widget(self.layout)
        self.add_widget(scroll)

    def abrir_dropdown(self, *args):
        self.dropdown.open()

    def set_tratamento(self, valor):
        self.tratamento = valor
        self.tratamento_botao.children[0].text = valor
        self.dropdown.dismiss()

    def proxima_tela(self, *args):
        print("Tratamento:", self.tratamento)
        print("Nome:", self.proponente.text)
        # ... continue com a lógica


class MainApp(MDApp):
    def build(self):
        self.title = "Tela de Dados Refinada"
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Dark"
        return DadosScreen()


if __name__ == "__main__":
    MainApp().run()
