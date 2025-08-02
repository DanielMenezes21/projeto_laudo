from kivymd.uix.screen import MDScreen
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText, MDTextFieldHelperText, MDTextFieldTrailingIcon
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivy.clock import Clock
import os
from datetime import datetime
from modules.data_folder import formatar_data
from kivy.uix.widget import Widget
from kivy.metrics import dp
from kivy.core.window import Window

data = formatar_data()
data_nome = datetime.now()
mes = f"{data_nome.month:02d}. {data.split('de')[1].strip()}"
class ImprovementScreen(MDScreen):
    def definir_matricula(self, indice, dados_matricula):
        self.indice_matricula = indice
        self.dados_matricula = dados_matricula
        print(f"📌 Benfeitoria vinculada à matrícula {indice}: {dados_matricula.get('nome_imovel', 'Imóvel sem nome')}")

    def __init__ (self, **kwargs):
        super().__init__(**kwargs)
        self.layout = MDBoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10), size_hint=(1, 1))
        self.add_widget(self.layout)

        self.scroll_view = MDScrollView(size_hint=(1, None), size=(Window.width, Window.height))
        self.layout.add_widget(self.scroll_view)

        self.content_box = MDBoxLayout(orientation='vertical', size_hint_y=None)
        self.content_box.bind(minimum_height=self.content_box.setter('height'))
        self.scroll_view.add_widget(self.content_box)

        self.title_label = MDLabel(
            text="Benfeitorias",
            halign="center",
            size_hint_y=None,
            height=dp(40)
        )
        self.content_box.add_widget(self.title_label)

        self.image_button = MDButton(
            MDButtonIcon(icon="image"),
            MDButtonText(text="Selecionar imagem"),
            size_hint=(1, None),
            on_release=lambda x: self.abrir_file_manager()
        )
        self.content_box.add_widget(self.image_button)

        self.largura_field = MDTextField(
            MDTextFieldHintText(text="Largura (m)"),
            size_hint=(1, None),
            height=dp(48)
        )
        self.comprimento_field = MDTextField(
            MDTextFieldHintText(text="Comprimento (m)"),
            size_hint=(1, None),
            height=dp(48)
        )
        self.altura_field = MDTextField(
            MDTextFieldHintText(text="Altura (m)"),
            size_hint=(1, None),
            height=dp(48)
        )
        self.area_field = MDTextField(
            MDTextFieldHintText(text="Área total (m²)"),
            size_hint=(1, None),
            height=dp(48)
        )

        for campo in [self.largura_field, self.comprimento_field, self.altura_field, self.area_field]:
            self.content_box.add_widget(campo)

        # Campo de descrição
        self.descricao_field = MDTextField(
            MDTextFieldHintText(text="Descrição da benfeitoria"),
            multiline=True,
            size_hint=(1, None),
            height=dp(100)
        )
        self.content_box.add_widget(self.descricao_field)

        # Checkboxes para estado de conservação
        self.estado_label = MDLabel(text="Estado de conservação:", size_hint_y=None, height=dp(30))
        self.content_box.add_widget(self.estado_label)

        self.checkbox_estado = {}
        for estado in ["Bom", "Mediano", "Ruim"]:
            box = MDBoxLayout(orientation="horizontal", size_hint=(1, None), height=dp(40), padding=(10, 0))
            chk = MDCheckbox(group="estado")
            label = MDLabel(text=estado, valign="center")
            box.add_widget(chk)
            box.add_widget(label)
            self.content_box.add_widget(box)
            self.checkbox_estado[estado.lower()] = chk

        self.botao_concluido = MDButton(
            MDButtonText(text="Concluído"),
            size_hint=(1, None),
            on_release=self.salvar_benfeitoria
        )
        self.content_box.add_widget(self.botao_concluido)

    def abrir_file_manager(self):
        initial_path = r"\\10.0.100.160\\Agropassos\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
        initial_path = os.path.join(initial_path, mes)
        if not os.path.exists(initial_path):
            initial_path = os.path.join(os.path.expanduser("~/Documents"))
        self.file_manager = MDFileManager(
            exit_manager=self.fechar_file_manager,
            select_path=self.selecionar_imagem,
            preview=True
        )
        self.file_manager.show(initial_path)  

    def fechar_file_manager(self, *args):
        self.file_manager.close()

    def selecionar_imagem(self, path):
        from kivy.uix.image import Image
        self.fechar_file_manager()
        self.caminho_imagem = path
        print(f"📷 Imagem selecionada: {path}")

        if hasattr(self, "imagem_widget") and self.imagem_widget in self.content_box.children:
            self.content_box.remove_widget(self.imagem_widget)

        self.imagem_widget = Image(
            source=path,
            size_hint=(1, None),
            height=dp(200),
            allow_stretch=True,
            keep_ratio=True
        )
        self.content_box.add_widget(self.imagem_widget, index=2)  # logo após o botão de imagem

    def salvar_benfeitoria(self, *args):
        if not hasattr(self, "indice_matricula") or not hasattr(self, "dados_matricula"):
            print("❌ Nenhuma matrícula definida.")
            return

        estado = ""
        for key, checkbox in self.checkbox_estado.items():
            if checkbox.active:
                estado = key
                break

        self.dados_matricula["descricao_benfeitoria"] = self.descricao_field.text
        self.dados_matricula["largura"] = self.largura_field.text
        self.dados_matricula["comprimento"] = self.comprimento_field.text
        self.dados_matricula["altura"] = self.altura_field.text
        self.dados_matricula["area"] = self.area_field.text
        self.dados_matricula["estado_benfeitoria"] = estado
        self.dados_matricula["imagem_benfeitoria"] = getattr(self, "caminho_imagem", "")

        print(f"✅ Benfeitoria salva para matrícula {self.dados_matricula.get('matricula', '')}")
        print(f"Dados da benfeitoria: {self.dados_matricula}")

        # Acessa a tela 'territorio' para obter a lista completa e imprimi-la
        tela_territorio = self.manager.get_screen('territorio')

        self.manager.current = "territorio"
