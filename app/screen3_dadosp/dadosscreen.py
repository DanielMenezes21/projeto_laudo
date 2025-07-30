from kivymd.uix.screen import MDScreen
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.dropdownitem import MDDropDownItem, MDDropDownItemText
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from app.screen3_dadosp.dados_function import receber_dados_pdf, go_back, go_next, fechar_arquivo, on_pdf_selecionado
from kivy.uix.widget import Widget
from kivy.metrics import dp
from kivy.clock import Clock
from app.screen3_dadosp.dados_function import abrir_dialogo_matriculas
import os
import re
from docx import Document

class DadosScreen(MDScreen):
    def open_dropdown(self, *args):
        self.dropdown.open()

    def set_tratamento(self, valor):
        self.tratamento = valor
        self.botao.children[0].text = valor
        self.dropdown.dismiss()
        if self.proprietario_atual in self.proprietarios:
            self.proprietarios[self.proprietario_atual]["tratamento"] = valor
    
    def salvar_civil(self, instance, value):
        if not value and self.proprietario_atual in self.proprietarios:
            self.proprietarios[self.proprietario_atual]["civil"] = instance.text

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.qtd_imoveis = 0
        self.tratamento = ""

        self.scroll = MDScrollView(bar_color=(1, 1, 1, 0.5), bar_width=10, scroll_type=["bars", "content"])

        self.layout = MDBoxLayout(orientation='vertical', padding=15, spacing=20, size_hint_y=None)
        self.layout.bind(minimum_height=self.layout.setter('height'))
        buttons = MDFloatLayout(size_hint_y=None)
        buttons.add_widget(Widget())
        cliente = MDBoxLayout(orientation="horizontal", spacing=10, size_hint_y=None, height=dp(48))

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_back(self),
        )
        buttons.add_widget(self.button_back)

        self.button_next = MDIconButton(
            icon="arrow-right",
            size_hint=(0.1, None),
            pos_hint={"x": 0.9, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_next(self),
        )
        buttons.add_widget(self.button_next)

        self.label = MDLabel(
            text="Dados do proprietario",
            halign="center",
            theme_text_color="Custom",
            text_color="yellow",
            size_hint=(1, None),
            pos_hint={"center_x": 0.5, "center_y":0.9}
        )

        self.solicitante = MDTextField(
            MDTextFieldHintText(text="Solicitante"),
            MDTextFieldHelperText(text="O solicitante que enviou os documentos para o laudo",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.botao = MDButton(
            MDButtonText(text="Selecionar Tratamento"),
            pos_hint={"center_x": 0.5},
            on_release=self.open_dropdown
        )
        cliente.add_widget(self.botao)
        
        self.proprietarios = {}
        self.proprietario_atual = ""
        self.proprietario = MDTextField(
            MDTextFieldHintText(text="Proprietario"),
            MDTextFieldHelperText(text="Nome do proprietário",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )
        cliente.add_widget(self.proprietario)

        self.botao_matricula = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: abrir_dialogo_matriculas(self)
        )
        self.botao_matricula.add_widget(MDButtonText(text="Definir Matrículas"))

        self.cpf = MDTextField(
            MDTextFieldHintText(text="CPF",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            write_tab=False,
            pos_hint={"center_x": 0.5},
            height=50,
        )

        self.civil = MDTextField(
            MDTextFieldHintText(text="situação civil"),
            MDTextFieldHelperText(text="fale sobre a situação civil do proprietario, se o mesmo se encontra casado,\n solteiro, viuvo ou se outra pessoa partilha a terra com o mesmo",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.municipio = MDTextField(
            MDTextFieldHintText(text="Município do imóvel"),
            MDTextFieldHelperText(text="Município onde o imóvel está localizado",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.estado = MDTextField(
            MDTextFieldHintText(text="Estado do imóvel"),
            MDTextFieldHelperText(text="Estado onde o imóvel está localizado",
                theme_text_color="Custom", 
                text_color_normal="yellow",
                text_color_focus="yellow",
                mode="on_focus",),
            theme_text_color="Custom",
            text_color_normal="yellow",
            text_color_focus="yellow",
            size = (200,50),
            size_hint=(0.9, None),
            pos_hint={"center_x": 0.5},
            write_tab=False,
            height=50,
        )

        self.botao_proprietario = MDButton(
            pos_hint={"center_x": 0.5},
            on_release=lambda x: self.menu_proprietario.open()
        )
        self.botao_proprietario.add_widget(MDButtonText(text="Selecionar proprietario"))

        self.menu_proprietario = MDDropdownMenu(
            caller=self.botao_proprietario,
            items=[],  
            width_mult=4,
        )
        cliente.add_widget(self.botao_proprietario)

        self.botao_acrescentar = MDIconButton(
            icon="plus",
            pos_hint={"center_x": 0.5},
            on_release=lambda x: self.acrescentar_proprietario()
        )

        self.dropdown = MDDropdownMenu(
            caller=self.botao,
            items=[
                {"text": "Sr.", "on_release": lambda x="Sr.": self.set_tratamento(x)},
                {"text": "Srª", "on_release": lambda x="Srª": self.set_tratamento(x)},
            ],
            width_mult=3
        )  

        self.layout.add_widget(buttons)
        self.layout.add_widget(self.label)
        self.layout.add_widget(self.solicitante)
        self.layout.add_widget(self.botao_matricula)
        self.layout.add_widget(self.civil)
        self.layout.add_widget(cliente)
        self.layout.add_widget(self.cpf)
        self.layout.add_widget(self.municipio)
        self.layout.add_widget(self.estado)
        self.layout.add_widget(self.botao_acrescentar)

        self.solicitante.bind(focus=self._on_focus)
        self.civil.bind(focus=self._on_focus)
        self.cpf.bind(focus=self._on_focus)
        self.municipio.bind(focus=self._on_focus)
        self.estado.bind(focus=self._on_focus)
        self.civil.bind(focus=self.salvar_civil)
        
        self.scroll.add_widget(self.layout)

        self.add_widget(self.scroll)

        Window.bind(on_keyboard_height=self._ajustar_scroll)
        self.campo_em_foco = None

    def _on_focus(self, instance, value):
        if value: 
            self.campo_em_foco = instance
            Clock.schedule_once(lambda dt: self.scroll.scroll_to(instance), 0.1)

    def _ajustar_scroll(self, window, altura_teclado):
        if altura_teclado > 0 and self.campo_em_foco:
            Clock.schedule_once(lambda dt: self.scroll.scroll_to(self.campo_em_foco), 0.1)
        else:
            self.scroll.scroll_y = 1

    def selecionar_proprietario(self, nome_escolhido):
        print(f"Selecionando: {nome_escolhido}")
        print(f"proprietarios: {self.proprietarios}")
        
        if self.proprietario_atual and self.proprietario_atual in self.proprietarios:
            self.proprietarios[self.proprietario_atual]["cpf"] = self.cpf.text.strip()
            self.proprietarios[self.proprietario_atual]["civil"] = self.civil.text.strip()
            self.proprietarios[self.proprietario_atual]["tratamento"] = self.tratamento

        self.proprietario_atual = nome_escolhido
        dados = self.proprietarios.get(self.proprietario_atual, {})
        print(f"Atualizando UI com: {dados}")
        
        self.proprietario.text = dados.get("nome", self.proprietario_atual)
        self.cpf.text = dados.get("cpf", "")
        self.civil.text = dados.get("civil", "")
        self.set_tratamento(dados.get("tratamento", ""))

        self.menu_proprietario.items = [
            {"text": nome, "on_release": lambda x=None, nome=nome: self.selecionar_proprietario(nome)}
            for nome in self.proprietarios
        ]
        self.menu_proprietario.dismiss()

    def set_novo_imovel(self, matricula):
        self.novo_imovel_selecionado = matricula
        self.novo_imovel_btn.children[0].text = matricula
        self.dropdown_imovel.dismiss()

    def acrescentar_proprietario(self, *args):
        self.novo_nome = MDTextField(
            MDTextFieldHintText(text="Nome do proprietário"),
            size_hint=(0.9, None),
            height=dp(50),
            pos_hint={"center_x": 0.5},
            write_tab=False
        )
        self.novo_cpf = MDTextField(
            MDTextFieldHintText(text="CPF"),
            size_hint=(0.9, None),
            height=dp(50),
            pos_hint={"center_x": 0.5},
            write_tab=False
        )
        self.novo_civil = MDTextField(
            MDTextFieldHintText(text="Situação Civil"),
            size_hint=(0.9, None),
            height=dp(50),
            pos_hint={"center_x": 0.5},
            write_tab=False
        )

        self.novo_tratamento = MDButton(
            MDButtonText(text="Selecionar Tratamento"),
            pos_hint={"center_x": 0.5}
        )
        self.dropdown_tratamento = MDDropdownMenu(
            caller=self.novo_tratamento,
            items=[
                {"text": "Sr.", "on_release": lambda x="Sr.": self.set_novo_tratamento(x)},
                {"text": "Srª", "on_release": lambda x="Srª": self.set_novo_tratamento(x)},
            ],
            width_mult=3
        )
        self.novo_tratamento.bind(on_release=lambda x: self.dropdown_tratamento.open())
        self.novo_tratamento_valor = ""

        self.novo_imovel_btn = MDButton(
            MDButtonText(text="Selecionar Imóvel/Matrícula"),
            pos_hint={"center_x": 0.5}
        )
        self.novo_imovel_selecionado = ""

        tela_matricula = self.manager.get_screen('matricula')
        matriculas = [
            f"{dados['nome_imovel']} - {dados['matricula']}"
            if dados['matricula'] else dados['nome_imovel']
            for dados in tela_matricula.lista_dados_matriculas
            if dados['nome_imovel']
        ]

        self.dropdown_imovel = MDDropdownMenu(
            caller=self.novo_imovel_btn,
            items=[
                {"text": matricula, "on_release": lambda x=matricula: self.set_novo_imovel(x)}
                for matricula in matriculas
            ],
            width_mult=4
        )
        self.novo_imovel_btn.bind(on_release=lambda x: self.dropdown_imovel.open())

        popup_layout = MDBoxLayout(
            orientation="vertical",
            padding=dp(10),
            spacing=dp(10),
            size_hint=(0.9, None),
            height=dp(350)
        )
        popup_layout.add_widget(self.novo_nome)
        popup_layout.add_widget(self.novo_imovel_btn)
        popup_layout.add_widget(self.novo_cpf)
        popup_layout.add_widget(self.novo_civil)
        popup_layout.add_widget(self.novo_tratamento)

        def salvar_novo_proprietario(instance):
            nome = self.novo_nome.text.strip()
            matricula = self.novo_imovel_selecionado

            if not nome or not matricula:
                print("❌ Nome do proprietário e matrícula não podem estar vazios.")
                return

            # Extrai nome_imovel e número da matrícula
            nome_imovel, matricula_num = (
                matricula.split(" - ", 1) if " - " in matricula else (matricula, "")
            )

            self.proprietarios[nome] = {
                "nome": nome,
                "cpf": self.novo_cpf.text.strip(),
                "civil": self.novo_civil.text.strip(),
                "tratamento": self.novo_tratamento_valor,
                "nome_imovel": nome_imovel,
                "matricula": matricula_num
            }

            self.menu_proprietario.items = [
                {"text": k, "on_release": lambda x=None, k=k: self.selecionar_proprietario(k)}
                for k in self.proprietarios
            ]

            # Atualiza dados_imoveis
            if not hasattr(self, 'dados_imoveis'):
                self.dados_imoveis = []
            self.dados_imoveis.append({
                "nome_imovel": nome_imovel,
                "matricula": matricula_num,
                "proprietario": nome,
                "latitude": "",
                "longitude": ""
            })

            print(f"✅ Proprietário {nome} adicionado com matrícula: {matricula}")
            self.dialogo_proprietario.dismiss()

        # Cria o popup
        self.dialogo_proprietario = MDDialog(
            MDDialogHeadlineText(text="Adicionar Proprietário"),
            popup_layout,
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="Cancelar"),
                    on_release=lambda x: self.dialogo_proprietario.dismiss()
                ),
                MDButton(
                    MDButtonText(text="Salvar"),
                    on_release=salvar_novo_proprietario
                )
            ),
            auto_dismiss=False
        )
        self.dialogo_proprietario.open()

    def set_novo_tratamento(self, valor):
        self.novo_tratamento_valor = valor
        self.novo_tratamento.children[0].text = valor
        self.dropdown_tratamento.dismiss()
