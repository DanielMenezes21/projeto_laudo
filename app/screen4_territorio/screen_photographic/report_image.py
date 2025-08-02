from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.dialog import MDDialog, MDDialogContentContainer, MDDialogButtonContainer, MDDialogHeadlineText
from kivy.uix.image import AsyncImage
from kivy.metrics import dp
from kivymd.app import MDApp
from datetime import datetime
from modules.data_folder import formatar_data
import os

data = formatar_data()
data_nome = datetime.now()
mes = f"{data_nome.month:02d}. {data.split('de')[1].strip()}"
class ReportImageScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.imagens_detalhadas = []

        self.layout = MDBoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10), size_hint=(1, 1))
        self.add_widget(self.layout)

        self.scroll = MDScrollView()
        self.content_box = MDBoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(10), padding=dp(10))
        self.content_box.bind(minimum_height=self.content_box.setter("height"))
        self.scroll.add_widget(self.content_box)
        self.layout.add_widget(self.scroll)

        self.botoes_box = MDBoxLayout(size_hint=(1, None), height=dp(56), spacing=dp(10))

        self.botao_adicionar = MDButton(
            MDButtonIcon(icon="image-plus"),
            MDButtonText(text="Adicionar imagem"),
            size_hint=(.5, None),
            on_release=self.abrir_file_manager
        )

        self.botao_concluir = MDButton(
            MDButtonIcon(icon="check"),
            MDButtonText(text="Concluir"),
            size_hint=(.5, None),
            on_release=self.finalizar_insercao
        )

        self.botoes_box.add_widget(self.botao_adicionar)
        self.botoes_box.add_widget(self.botao_concluir)
        self.layout.add_widget(self.botoes_box)

    def atualizar_info(self, imagem_path, nova_desc, novo_tipo, campos_dimensao):
        for item in self.imagens_detalhadas:
            if item["imagem"] == imagem_path:
                item["descricao"] = nova_desc
                item["tipo"] = novo_tipo
                if novo_tipo.lower() == "estrutura":
                    for i, chave in enumerate(["largura", "comprimento", "altura", "area"]):
                        item[chave] = campos_dimensao[i].text
                break

    def abrir_file_manager(self, *args):
        initial_path = r"\\10.0.100.160\\Agropassos\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
        initial_path = os.path.join(initial_path, mes)
        if not os.path.exists(initial_path):
            initial_path = os.path.join(os.path.expanduser("~/Documents"))
        self.file_manager = MDFileManager(
            exit_manager=self.fechar_file_manager,
            select_path=self.adicionar_imagem,
            preview=True
        )
        self.file_manager.show(initial_path)

    def fechar_file_manager(self, *args):
        if hasattr(self, "file_manager"):
            self.file_manager.close()

    def adicionar_imagem(self, path):
        self.fechar_file_manager()
        if not any(d["imagem"] == path for d in self.imagens_detalhadas):
            self.mostrar_formulario_imagem(path)

    def mostrar_formulario_imagem(self, imagem_path):
        self.descricao = MDTextField(
            MDTextFieldHintText(text="Descrição da imagem"),
            multiline=True,
            size_hint=(1, None),
            height=dp(100)
        )

        self.tipo_imagem = "lavoura"
        self.botao_lavoura = MDButton(
            MDButtonText(text="Lavoura"),
            on_release=lambda x: self.toggle_campos_dimensao("lavoura"),
            size_hint=(.5, None),
            height=dp(40)
        )

        self.botao_estrutura = MDButton(
            MDButtonText(text="Estrutura"),
            on_release=lambda x: self.toggle_campos_dimensao("estrutura"),
            size_hint=(.5, None),
            height=dp(40)
        )

        self.tipo_buttons_box = MDBoxLayout(orientation="horizontal", spacing=dp(10), size_hint=(1, None), height=dp(40))
        self.tipo_buttons_box.add_widget(self.botao_lavoura)
        self.tipo_buttons_box.add_widget(self.botao_estrutura)

        self.tipo_imagem = "lavoura"

        self.dimensoes_box = MDBoxLayout(orientation="vertical", size_hint_y=None, height=0)
        self.largura = MDTextField(MDTextFieldHintText(text="Largura (m)"), size_hint=(1, None), height=dp(48), opacity=0)
        self.comprimento = MDTextField(MDTextFieldHintText(text="Comprimento (m)"), size_hint=(1, None), height=dp(48), opacity=0)
        self.altura = MDTextField(MDTextFieldHintText(text="Altura (m)"), size_hint=(1, None), height=dp(48), opacity=0)
        self.area = MDTextField(MDTextFieldHintText(text="Área (m²)"), size_hint=(1, None), height=dp(48), opacity=0)

        for campo in [self.largura, self.comprimento, self.altura, self.area]:
            self.dimensoes_box.add_widget(campo)

        content = MDBoxLayout(orientation="vertical", spacing=10, size_hint_y=None)
        content.add_widget(self.tipo_buttons_box)
        content.add_widget(self.descricao)
        content.add_widget(self.dimensoes_box)

        btn_salvar = MDButton(
            MDButtonText(text="Salvar"),
            on_release=lambda x: self.salvar_info_imagem(imagem_path)
        )
        btn_cancelar = MDButton(
            MDButtonText(text="Cancelar"),
            on_release=lambda x: self.form_dialog.dismiss()
        )

        self.form_dialog = MDDialog(
            MDDialogHeadlineText(text="Informações da Imagem"),
            MDDialogContentContainer(content),
            MDDialogButtonContainer(btn_cancelar, btn_salvar)
        )
        self.form_dialog.open()

    def toggle_campos_dimensao(self, tipo):
        self.tipo_imagem = tipo
        if tipo == "estrutura":
            self.dimensoes_box.height = dp(48 * 4)
            for campo in self.dimensoes_box.children:
                campo.opacity = 1
                campo.disabled = False
            self.botao_estrutura.style = "filled"
            self.botao_lavoura.style = "outlined"
        else:
            self.dimensoes_box.height = 0
            for campo in self.dimensoes_box.children:
                campo.opacity = 0
                campo.disabled = True
            self.botao_lavoura.style = "filled"
            self.botao_estrutura.style = "outlined"

    def salvar_info_imagem(self, imagem_path):
        dados = {
            "imagem": imagem_path,
            "tipo": self.tipo_imagem,
            "descricao": self.descricao.text.strip()
        }

        if self.tipo_imagem == "estrutura":
            dados.update({
                "largura": self.largura.text,
                "comprimento": self.comprimento.text,
                "altura": self.altura.text,
                "area": self.area.text
            })

        self.imagens_detalhadas.append(dados)
        img_box = MDBoxLayout(orientation="horizontal", spacing=dp(10), size_hint_y=None, height=dp(200))

        imagem_widget = AsyncImage(source=imagem_path, size_hint=(None, 1), width=dp(160), allow_stretch=True)
        img_box.add_widget(imagem_widget)

        info_box = MDBoxLayout(orientation="vertical", spacing=dp(5), size_hint_y=1)

        descricao_field = MDTextField(
            MDTextFieldHintText(text="Descrição"),
            text=dados.get("descricao", ""),
            size_hint=(1, None),
            height=dp(48)
        )

        tipo_field = MDTextField(
            MDTextFieldHintText(text="Tipo (estrutura/lavoura)"),
            text=dados.get("tipo", ""),
            size_hint=(1, None),
            height=dp(48)
        )

        dim_fields = []
        if dados.get("tipo", "") == "estrutura":
            for campo_nome in ["largura", "comprimento", "altura", "area"]:
                dim_fields.append(MDTextField(
                    MDTextFieldHintText(text=f"{campo_nome.capitalize()} (m)"),
                    text=dados.get(campo_nome, ""),
                    size_hint=(1, None),
                    height=dp(48)
                ))

        btn_salvar = MDButton(
            MDButtonText(text="Salvar"),
            size_hint=(1, None),
            on_release=lambda x, desc=descricao_field, tipo=tipo_field, dims=dim_fields: self.atualizar_info(
                imagem_path, desc.text, tipo.text, dims
            )
        )

        info_box.add_widget(descricao_field)
        info_box.add_widget(tipo_field)
        for field in dim_fields:
            info_box.add_widget(field)
        info_box.add_widget(btn_salvar)

        img_box.add_widget(info_box)
        self.content_box.add_widget(img_box)

        self.form_dialog.dismiss()

    def finalizar_insercao(self, *args):
        if hasattr(self.manager, "callback_fotos"):
            self.manager.callback_fotos(self.imagens_detalhadas)
        self.manager.current = "territorio"
