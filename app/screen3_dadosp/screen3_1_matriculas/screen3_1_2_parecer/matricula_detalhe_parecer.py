from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.button import MDIconButton
from kivymd.uix.scrollview import MDScrollView
from kivy.uix.widget import Widget
from kivymd.uix.label import MDLabel
from kivy.metrics import dp
from kivy.properties import BooleanProperty, DictProperty, StringProperty

class MatriculaParecerScreen(MDScreen):
    biomas = {
        "Amazônia": False,
        "Cerrado": False,
        "Caatinga": False,
        "Mata Atlântica": False,
        "Pampa": False,
        "Pantanal": False
    }
    biomas_selecionados = DictProperty({})

    tipos_passivo = DictProperty({
        "Embargo": False,
        "Déficit de Reserva Legal": False,
        "Alerta MapBiomas": False
    })

    def voltar(self, *args):
        self.manager.current = 'matricula'

    def _carregar_dados_iniciais(self):
        """Preenche os campos com os dados iniciais da matrícula"""
        if not hasattr(self, 'dados_matricula'):
            return

        self.campo_texto.text = self.dados_matricula.get('observacoes_parecer', '')
        self.campo_car.text = self.dados_matricula.get('car', '')
        
        self.georreferenciamento_sim = self.dados_matricula.get('possui_georref', False)
        self.campo_georref.text = self.dados_matricula.get('numero_georref', '')
        self.campo_georref.disabled = not self.georreferenciamento_sim
        
        self.alienacao_sim = self.dados_matricula.get('possui_alienacao', False)
        self.campo_alienacao.text = self.dados_matricula.get('detalhes_alienacao', '')
        self.campo_alienacao.disabled = not self.alienacao_sim
        
        self.apa_sim = self.dados_matricula.get('possui_apa', False)
        self.campo_apa.text = self.dados_matricula.get('nome_apa', '')
        self.campo_apa.disabled = not self.apa_sim
        
        biomas_salvos = self.dados_matricula.get('biomas', {})
        for bioma in self.biomas_selecionados:
            self.biomas_selecionados[bioma] = biomas_salvos.get(bioma, False)
        
        self.passivo_ambiental_sim = self.dados_matricula.get('possui_passivo', False)
        self.layout_tipos_passivo.disabled = not self.passivo_ambiental_sim
        self.campo_detalhes_passivo.disabled = not self.passivo_ambiental_sim
        
        tipos_passivo_salvos = self.dados_matricula.get('tipos_passivo', {})
        for tipo, estado in tipos_passivo_salvos.items():
            if tipo in self.tipos_passivo:
                self.tipos_passivo[tipo] = estado
        
        self.campo_detalhes_passivo.text = self.dados_matricula.get('detalhes_passivo', '')

    def __init__(self, nome_matricula, detalhes_screen=None, lista_dados_matriculas=None, indice_matricula=None, **kwargs):
        super().__init__(**kwargs)
        self.nome_matricula = nome_matricula
        self.detalhes_screen = detalhes_screen
        self.lista_dados_matriculas = lista_dados_matriculas or []
        self.indice_matricula = indice_matricula
        
        if lista_dados_matriculas is None:
            lista_dados_matriculas = []
        self.lista_dados_matriculas = lista_dados_matriculas
        
        if indice_matricula is None or not (0 <= indice_matricula < len(lista_dados_matriculas)):
            self.indice_matricula = len(lista_dados_matriculas) - 1 if lista_dados_matriculas else 0
        else:
            self.indice_matricula = indice_matricula
        
        self.dados_matricula = self.lista_dados_matriculas[self.indice_matricula] if self.lista_dados_matriculas else {}
        
        self.georreferenciamento_sim = BooleanProperty(False)
        self.alienacao_sim = BooleanProperty(False)
        self.apa_sim = BooleanProperty(False)
        self.passivo_ambiental_sim = BooleanProperty(False)
        
        self.biomas_selecionados = self.biomas.copy()
        
        self.detalhes_passivo = StringProperty("")

        kwargs.pop('detalhes_screen', None)
        kwargs.pop('lista_dados_matriculas', None)
        kwargs.pop('indice_matricula', None)
        super().__init__(**kwargs)

        layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)
        scroll = MDScrollView()
        scroll_content = MDBoxLayout(orientation="vertical", spacing=20, padding=10, size_hint_y=None)
        scroll_content.bind(minimum_height=scroll_content.setter('height'))

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=self.voltar,
        )
        scroll_content.add_widget(self.button_back)

        self.campo_texto = MDTextField(
            MDTextFieldHintText(text=f"Observações - {nome_matricula}"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5}
        )
        scroll_content.add_widget(self.campo_texto)

        self.campo_car = MDTextField(
            MDTextFieldHintText(text=f"CAR - {nome_matricula}"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5}
        )
        scroll_content.add_widget(self.campo_car)

        scroll_content.add_widget(self._criar_selecao_sim_nao(
            "Possui georreferenciamento?",
            "georref",
            lambda x: self._ativar_campo(self.campo_georref, x, 'georreferenciamento_sim')
        ))
        self.campo_georref = MDTextField(
            MDTextFieldHintText(text="Número do georreferenciamento"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5},
            disabled=True
        )
        scroll_content.add_widget(self.campo_georref)

        scroll_content.add_widget(self._criar_selecao_sim_nao(
            "Possui alienação fiduciária?",
            "alienacao",
            lambda x: self._ativar_campo(self.campo_alienacao, x, 'alienacao_sim')
        ))
        self.campo_alienacao = MDTextField(
            MDTextFieldHintText(text="Detalhes da alienação fiduciária"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5},
            disabled=True
        )
        scroll_content.add_widget(self.campo_alienacao)

        scroll_content.add_widget(self._criar_selecao_sim_nao(
            "O imóvel está inserido em Área de Proteção Ambiental - APA?",
            "apa",
            lambda x: self._ativar_campo(self.campo_apa, x, 'apa_sim')
        ))
        self.campo_apa = MDTextField(
            MDTextFieldHintText(text="Nome da APA"),
            size_hint=(0.9, None),
            height=50,
            pos_hint={"center_x": 0.5},
            disabled=True
        )
        scroll_content.add_widget(self.campo_apa)

        layout_biomas = MDBoxLayout(orientation="vertical", spacing=10, size_hint=(0.9, None), pos_hint={"center_x": 0.5})
        label_biomas = MDLabel(text="Biomas do imóvel:", size_hint_y=None, height=30)
        layout_biomas.add_widget(label_biomas)
        
        linha1 = MDBoxLayout(orientation="horizontal", spacing=20, size_hint_y=None, height=40)
        for bioma in ["Amazônia", "Cerrado", "Caatinga"]:
            linha1.add_widget(self._criar_checkbox_bioma(bioma))
        
        linha2 = MDBoxLayout(orientation="horizontal", spacing=20, size_hint_y=None, height=40)
        for bioma in ["Mata Atlântica", "Pampa", "Pantanal"]:
            linha2.add_widget(self._criar_checkbox_bioma(bioma))
        
        layout_biomas.add_widget(linha1)
        layout_biomas.add_widget(linha2)
        scroll_content.add_widget(layout_biomas)

        scroll_content.add_widget(self._criar_selecao_sim_nao(
            "Possui Passivo Ambiental?",
            "passivo",
            lambda x: self._ativar_secao_passivo(x)
        ))
        
        self.layout_tipos_passivo = MDBoxLayout(
            orientation="vertical", 
            spacing=10, 
            size_hint=(0.9, None),
            height=150,
            pos_hint={"center_x": 0.5},
            disabled=True
        )
        
        self.checkboxes_passivo = []
        for tipo in ["Embargo", "Déficit de Reserva Legal", "Alerta MapBiomas"]:
            box = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=40)
            checkbox = MDCheckbox(
                size_hint_x=None,
                width=30,
                on_release=lambda x, t=tipo: self._atualizar_tipo_passivo(t, x.active)
            )
            self.checkboxes_passivo.append(checkbox)
            box.add_widget(checkbox)
            box.add_widget(MDLabel(text=tipo, size_hint_x=None, width=200))
            self.layout_tipos_passivo.add_widget(box)
        scroll_content.add_widget(self.layout_tipos_passivo)
        
        self.campo_detalhes_passivo = MDTextField(
            MDTextFieldHintText(text="Detalhamento do passivo ambiental"),
            size_hint=(0.9, None),
            height=100,
            pos_hint={"center_x": 0.5},
            disabled=True,
            multiline=True
        )
        scroll_content.add_widget(self.campo_detalhes_passivo)

        self.button_concluido = MDIconButton(
            icon="check",
            size_hint=(1, None),
            pos_hint={"right": 1, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(0, 1, 0, 1),
            on_release=lambda x: self.salvar_dados(),
        )
        scroll_content.add_widget(self.button_concluido)
        
        scroll_content.add_widget(Widget(size_hint_y=None, height=20))
        
        scroll.add_widget(scroll_content)
        layout.add_widget(scroll)
        self.add_widget(layout)

        self._carregar_dados_iniciais()

    def _criar_checkbox_bioma(self, nome_bioma):
        """Creates a checkbox widget for a biome."""
        box = MDBoxLayout(orientation="horizontal", spacing=5, size_hint_x=None, width=150)
        checkbox = MDCheckbox(
            size_hint_x=None,
            width=30,
            active=self.biomas_selecionados.get(nome_bioma, False),
            on_release=lambda x, b=nome_bioma: self._atualizar_bioma(b, x.active)
        )
        box.add_widget(checkbox)
        box.add_widget(MDLabel(text=nome_bioma, size_hint_x=None, width=100))
        return box

    def _atualizar_bioma(self, bioma, estado):
        """Updates biome selection state."""
        current = dict(self.biomas_selecionados)
        current[bioma] = estado
        self.biomas_selecionados = current

    def _criar_selecao_sim_nao(self, texto, grupo, callback):
        layout = MDBoxLayout(orientation="horizontal", spacing=10, size_hint=(0.9, None), height=50, pos_hint={"center_x": 0.5})
        label = MDLabel(text=texto, size_hint_x=0.6, halign="left")

        checkbox_sim = MDCheckbox(
            size_hint_x=0.2,
            group=grupo,
            on_release=lambda chk: callback(True if chk.active else False)
        )

        checkbox_nao = MDCheckbox(
            size_hint_x=0.2,
            group=grupo,
            on_release=lambda chk: callback(False)
        )

        layout.add_widget(label)
        layout.add_widget(checkbox_sim)
        layout.add_widget(MDLabel(text="Sim", size_hint_x=0.1))
        layout.add_widget(checkbox_nao)
        layout.add_widget(MDLabel(text="Não", size_hint_x=0.1))
        return layout

    def _ativar_campo(self, campo, ativar, propriedade):
        """Activates/deactivates a conditional field."""
        campo.disabled = not ativar
        setattr(self, propriedade, ativar)

    def _ativar_secao_passivo(self, ativar):
        """Activates/deactivates the entire environmental liability section."""
        self.passivo_ambiental_sim = ativar
        self.layout_tipos_passivo.disabled = not ativar
        self.campo_detalhes_passivo.disabled = not ativar
        
        if not ativar:
            for checkbox in self.checkboxes_passivo:
                checkbox.active = False
            for tipo in self.tipos_passivo.keys():
                self.tipos_passivo[tipo] = False
            self.campo_detalhes_passivo.text = ""

    def _atualizar_tipo_passivo(self, tipo, estado):
        """Updates environmental liability type selection."""
        current = dict(self.tipos_passivo)
        current[tipo] = estado
        self.tipos_passivo[tipo] = current

    def salvar_dados(self):
        """Saves data without overwriting other matricula's data"""
        try:
            dados_atualizados = self.lista_dados_matriculas[self.indice_matricula].copy()
            
            campos_parecer = {
                "observacoes_parecer": self.campo_texto.text,
                "car": self.campo_car.text,
                "possui_georref": self.georreferenciamento_sim,
                "numero_georref": self.campo_georref.text if self.georreferenciamento_sim else "",
                "possui_alienacao": self.alienacao_sim,
                "detalhes_alienacao": self.campo_alienacao.text if self.alienacao_sim else "",
                "possui_apa": self.apa_sim,
                "nome_apa": self.campo_apa.text if self.apa_sim else "",
                "biomas": {k: v for k, v in self.biomas_selecionados.items() if v},
                "possui_passivo": self.passivo_ambiental_sim,
                "tipos_passivo": {k: v for k, v in self.tipos_passivo.items() if v},
                "detalhes_passivo": self.campo_detalhes_passivo.text if self.passivo_ambiental_sim else ""
            }
            if self.detalhes_screen:
                from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.detalhes_function import coletar_dados
                dados_detalhes = coletar_dados(self.detalhes_screen)
                for campo, valor in dados_detalhes.items():
                    dados_atualizados[campo] = valor
            
            for campo, valor in campos_parecer.items():
                dados_atualizados[campo] = valor

            self.lista_dados_matriculas[self.indice_matricula] = dados_atualizados
            
            print(f"✅ Dados salvos para matrícula {self.indice_matricula}")
            print(f"Total de matrículas: {len(self.lista_dados_matriculas)}")
            
        except Exception as e:
            print(f"❌ Erro ao salvar: {str(e)}")
            import traceback
            traceback.print_exc()