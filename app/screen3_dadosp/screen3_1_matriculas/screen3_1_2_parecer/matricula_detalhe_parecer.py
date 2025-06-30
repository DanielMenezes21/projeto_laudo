from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.button import MDIconButton
from kivymd.uix.label import MDLabel
from kivy.metrics import dp

class MatriculaParecerScreen(MDScreen):
    def voltar(self, *args):
        self.manager.current = 'matricula'

    def __init__(self, nome_matricula, detalhes_screen=None, lista_dados_matriculas=None, indice_matricula_atual=None, **kwargs):
        self.detalhes_screen = detalhes_screen
        self.lista_dados_matriculas = lista_dados_matriculas
        self.indice_matricula_atual = indice_matricula_atual
        self.nome_matricula = nome_matricula

        kwargs.pop('detalhes_screen', None)
        kwargs.pop('lista_dados_matriculas', None)
        kwargs.pop('indice_matricula_atual', None)
        super().__init__(**kwargs)
        # ... resto do código ...
        layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=self.voltar,
        )

        self.campo_texto = MDTextField(
            MDTextFieldHintText(text=f"Observações - {nome_matricula}"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5}
        )

        self.button_concluido = MDIconButton(
            icon="check",
            size_hint=(1, None),
            pos_hint={"right": 1, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(0, 1, 0, 1),  # verde
            on_release=lambda x: self.salvar_dados(),
        )

        layout.add_widget(self.button_back)
        layout.add_widget(self.campo_texto)
        layout.add_widget(self.button_concluido)

        self.add_widget(layout)
        
    def salvar_dados(self):
        dados = {
            "observacoes_parecer": self.campo_texto.text,
        }
        if self.detalhes_screen:
            from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.detalhes_function import coletar_checkboxes
            dados_det = coletar_checkboxes(self.detalhes_screen)
            dados.update(dados_det)
        if (
            self.lista_dados_matriculas
            and self.indice_matricula_atual is not None
            and 0 <= self.indice_matricula_atual < len(self.lista_dados_matriculas)
        ):
            self.lista_dados_matriculas[self.indice_matricula_atual].update(dados)
        else:
            print("❌ Índice fora do range ou lista vazia!")
        print("DEBUG após salvar:", self.lista_dados_matriculas)