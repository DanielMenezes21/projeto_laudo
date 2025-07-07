from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.button import MDIconButton
from kivymd.uix.label import MDLabel
from kivy.metrics import dp
from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.detalhes_function import go_back, coletar_checkboxes, ir_para_parecer

class MatriculaDetalheScreen(MDScreen):
    def coletar_detalhes(self):
        return coletar_checkboxes(self)
    
    def __init__(self, nome_matricula, lista_dados_matriculas=None, indice_matricula_atual=None, **kwargs):
        self.lista_dados_matriculas = lista_dados_matriculas
        self.indice_matricula_atual = indice_matricula_atual
        self.nome_matricula = nome_matricula
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

        self.campo_observacoes = MDTextField(
            MDTextFieldHintText(text="Observações sobre o imóvel"),
            multiline=True,
            size_hint=(0.9, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )

        linha_poligonos = MDBoxLayout(orientation="horizontal", spacing=30, size_hint_y=None, height=40)
        self.checkbox_regular = MDCheckbox()
        label_regular = MDLabel(text="Polígono regular", halign="left")
        self.checkbox_irregular = MDCheckbox()
        label_irregular = MDLabel(text="Polígono irregular", halign="left")
        linha_poligonos.add_widget(self.checkbox_regular)
        linha_poligonos.add_widget(label_regular)
        linha_poligonos.add_widget(self.checkbox_irregular)
        linha_poligonos.add_widget(label_irregular)

        label_declividade = MDLabel(
            text="Declividade",
            size_hint_x=None,
            width=dp(100),
        )

        linha_af = MDBoxLayout(orientation="horizontal", spacing=20, size_hint_y=None, height=40)
        self.checkboxes_af = {}
        for letra in ['A', 'B', 'C', 'D', 'E', 'F']:
            cb = MDCheckbox()
            lbl = MDLabel(text=letra, halign="left")
            self.checkboxes_af[letra] = cb
            box = MDBoxLayout(orientation="vertical", size_hint_x=None, width=40)
            box.add_widget(cb)
            box.add_widget(lbl)
            linha_af.add_widget(box)

        linha_pares = MDBoxLayout(orientation="horizontal", spacing=20, size_hint_y=None, height=40)
        self.checkboxes_pares = {}
        for par in ['AB', 'BA', 'BC', 'CB', 'CD', 'DC']:
            cb = MDCheckbox()
            lbl = MDLabel(text=par, halign="left")
            self.checkboxes_pares[par] = cb
            box = MDBoxLayout(orientation="vertical", size_hint_x=None, width=50)
            box.add_widget(cb)
            box.add_widget(lbl)
            linha_pares.add_widget(box)

        label_superficie = MDLabel(
            text="Superfície",
            size_hint_x=None,
            width=dp(100),
        )

        linha_superficie = MDBoxLayout(orientation="horizontal", spacing=20, size_hint_y=None, height=40)
        self.checkboxes_superficie = {}
        for superficie in ['seco', 'umido', 'alagadiço']:
            cb = MDCheckbox()
            lbl = MDLabel(text=superficie, halign="left")
            self.checkboxes_superficie[superficie] = cb
            box = MDBoxLayout(orientation="vertical", size_hint_x=None, width=80)
            box.add_widget(cb)
            box.add_widget(lbl)
            linha_superficie.add_widget(box)

        self.button_next = MDIconButton(
            icon="arrow-right",
            size_hint=(1, None),
            pos_hint={"right": 1, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: ir_para_parecer(self),
        )

        layout.add_widget(self.button_back)
        layout.add_widget(self.campo_observacoes)
        layout.add_widget(linha_poligonos)
        layout.add_widget(label_declividade)
        layout.add_widget(linha_af)
        layout.add_widget(linha_pares)
        layout.add_widget(label_superficie)
        layout.add_widget(linha_superficie)
        layout.add_widget(self.button_next)

        self.add_widget(layout)

        