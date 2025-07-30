from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.button import MDIconButton
from kivymd.uix.label import MDLabel
from kivymd.uix.scrollview import MDScrollView
from kivy.metrics import dp
import os
from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.detalhes_function import go_back, coletar_dados, ir_para_parecer, selecionar_imagem

class MatriculaDetalheScreen(MDScreen):
    def coletar_detalhes(self):
        return coletar_dados(self)
    
    def __init__(self, nome_matricula, lista_dados_matriculas=None, indice_matricula=None, **kwargs):
        self.lista_dados_matriculas = lista_dados_matriculas
        self.indice_matricula = indice_matricula
        self.nome_matricula = nome_matricula
        super().__init__(**kwargs)
        self.root = r"\\10.0.100.160\\Agropassos\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL"
        if not os.path.exists(self.root):
            self.root = os.path.join(os.path.expanduser("~/Documents"))
        self.current_path = self.root

        layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)
        scroll_view = MDScrollView()
        layout.size_hint_y = None
        layout.bind(minimum_height=layout.setter("height"))
        campo_reserva = MDBoxLayout(orientation="horizontal", padding=20, spacing=20, size_hint_y=None, height=dp(120))
        campo_app = MDBoxLayout(orientation="horizontal", padding=20, spacing=20, size_hint_y=None, height=dp(120))

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
            MDTextFieldHintText(text="Observações sobre o imóvel para o item 6.4 - descrição do imóvel"),
            multiline=True,
            size_hint=(0.9, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )

        self.p_reserva = MDTextField(
            MDTextFieldHintText(text="Porcentagem da área de reserva"),
            size_hint=(0.4, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )
        campo_reserva.add_widget(self.p_reserva)

        self.area_reserva = MDTextField(
            MDTextFieldHintText(text="Tamanho da área de reserva"),
            size_hint=(0.4, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )
        campo_reserva.add_widget(self.area_reserva)

        self.p_app = MDTextField(
            MDTextFieldHintText(text="Porcentagem da área de APP"),
            size_hint=(0.4, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )
        campo_app.add_widget(self.p_app)

        self.a_app = MDTextField(
            MDTextFieldHintText(text="Tamannho da área de APP"),
            size_hint=(0.4, None),
            height=100,
            pos_hint={"center_x": 0.5}
        )
        campo_app.add_widget(self.a_app)

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

        self.atividade = MDLabel(
            text=f"Atividade: {self.nome_matricula}",
            size_hint_y=None,
            height=dp(40),
            pos_hint={"center_x": 0.5}
        )

        self.atividade_imovel = MDTextField(
            MDTextFieldHintText(text="Atividade do imóvel"),
            size_hint=(0.9, None),
            height=40,
            pos_hint={"center_x": 0.5}
        )

        self.atividade_potencial = MDTextField(
            MDTextFieldHintText(text="Atividade potencial do imóvel"),
            size_hint=(0.9, None),
            height=40,
            pos_hint={"center_x": 0.5}
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

        self.satelite = MDLabel(
            text="Imagens de satélite",
            size_hint_x=None,
            width=dp(100),
        )

        selecionar_img = MDBoxLayout(orientation="horizontal", size_hint_y=None, height=dp(40))
        description_img = MDBoxLayout(orientation="horizontal", padding=20, spacing=20, size_hint_y=None, height=dp(120))

        self.image_one_selection = MDIconButton(
            icon="image",
            size_hint=(1, None),
            pos_hint={"center_x": 0.5},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: selecionar_imagem(self, self.image_one_selection),
        )
        selecionar_img.add_widget(self.image_one_selection)

        self.image_two_selection = MDIconButton(
            icon="image",
            size_hint=(1, None),
            pos_hint={"center_x": 0.5},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: selecionar_imagem(self, self.image_two_selection),
        )
        selecionar_img.add_widget(self.image_two_selection)

        self.description_img_one = MDTextField(
            MDTextFieldHintText(text="Descrição da imagem 1"),
            size_hint=(0.4, None),
            height=40,
            pos_hint={"center_x": 0.5}
        )
        description_img.add_widget(self.description_img_one)

        self.description_img_two = MDTextField(
            MDTextFieldHintText(text="Descrição da imagem 2"),
            size_hint=(0.4, None),
            height=40,
            pos_hint={"center_x": 0.5}
        )
        description_img.add_widget(self.description_img_two)

        self.button_next = MDIconButton(
            icon="arrow-right",
            size_hint=(1, None),
            pos_hint={"right": 1, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: ir_para_parecer(self),
        )

        self.img_one = ""
        self.img_two = ""

        layout.add_widget(self.button_back)
        layout.add_widget(self.campo_observacoes)
        layout.add_widget(campo_reserva)
        layout.add_widget(campo_app)
        layout.add_widget(linha_poligonos)
        layout.add_widget(label_declividade)
        layout.add_widget(linha_af)
        layout.add_widget(linha_pares)
        layout.add_widget(label_superficie)
        layout.add_widget(linha_superficie)
        layout.add_widget(self.atividade)
        layout.add_widget(self.atividade_imovel)
        layout.add_widget(self.atividade_potencial)
        layout.add_widget(self.satelite)
        layout.add_widget(selecionar_img)
        layout.add_widget(description_img)
        layout.add_widget(self.button_next)

        scroll_view.add_widget(layout)
        self.add_widget(scroll_view)

        