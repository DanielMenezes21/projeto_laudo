from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.button import MDIconButton
from kivymd.uix.label import MDLabel
from kivy.metrics import dp
from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.detalhes_function import go_back

class MatriculaDetalheScreen(MDScreen):
    def __init__(self, nome_matricula, **kwargs):
        super().__init__(**kwargs)

        layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_back(self),
        )

        self.campo_texto = MDTextField(
            hint_text=f"Observações - {nome_matricula}",
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5}
        )

        linha_checkbox = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=40)
        self.checkbox = MDCheckbox()
        label = MDLabel(text="Confirmado", halign="left")

        linha_checkbox.add_widget(self.checkbox)
        linha_checkbox.add_widget(label)

        layout.add_widget(self.button_back)
        layout.add_widget(self.campo_texto)
        layout.add_widget(linha_checkbox)

        self.add_widget(layout)
