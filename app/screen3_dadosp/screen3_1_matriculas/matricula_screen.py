from kivymd.uix.screen import MDScreen
from kivymd.uix.textfield import MDTextField
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivy.uix.widget import Widget
from datetime import datetime
from modules.data_folder import formatar_data
import os
from kivy.metrics import dp
from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.matricula_detalhe_screen import MatriculaDetalheScreen

class MatriculaScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)
        self.scroll = MDScrollView()
        self.botoes_matriculas = MDBoxLayout(orientation="vertical", spacing=10, size_hint_y=None)
        self.botoes_matriculas.bind(minimum_height=self.botoes_matriculas.setter("height"))
        self.scroll.add_widget(self.botoes_matriculas)

        self.layout.add_widget(self.scroll)
        self.add_widget(self.layout)

        self.file_manager_excel = MDFileManager(
            exit_manager=self.fechar_filemanager,
            select_path=self.selecionar_arquivo,
            ext=[".xlsx", ".xls"],
            preview=False
        )

        self.file_manager_imagem = MDFileManager(
            exit_manager=self.fechar_filemanager,
            select_path=self.selecionar_arquivo,
            ext=[".png", ".jpg", ".jpeg"],
            preview=True  
        )

        self.matriculas = {}
        self.matricula_selecionada = ""
        self.tipo_planilha = ""

    def abrir_filemanager(self, matricula_nome, tipo):
        data = formatar_data()
        data_nome = datetime.now()
        mes = f'{data_nome.month:02d}. {data.split("de")[1].strip()}'
        self.matricula_selecionada = matricula_nome
        self.tipo_planilha = tipo

        initial_path = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL"
        initial_path = os.path.join(initial_path, mes)
        if not os.path.exists(initial_path):
            initial_path = os.path.join(os.path.expanduser("~"), "Documents")
        self.root_path = initial_path
        self.current_path = self.root_path

        if tipo == "imagem":
            self.file_manager_imagem.show(initial_path)
        else:
            self.file_manager_excel.show(initial_path)


    def fechar_filemanager(self, *args):
        if getattr(self.file_manager_excel, "_window_manager", None):
            self.file_manager_excel.close()
        if getattr(self.file_manager_imagem, "_window_manager", None):
            self.file_manager_imagem.close()

    def selecionar_arquivo(self, caminho):
        self.fechar_filemanager()
        m = self.matricula_selecionada
        t = self.tipo_planilha
        self.matriculas[m]["arquivos"][t] = caminho
        print(f"📂 {t.upper()} associado a {m}: {caminho}")

    def abrir_tela_detalhe(self, nome_matricula):
        nome_tela = f"detalhe_{nome_matricula.replace(' ', '_').lower()}"
        if not self.manager.has_screen(nome_tela):
            nova_tela = MatriculaDetalheScreen(nome_matricula, name=nome_tela)
            self.manager.add_widget(nova_tela)
        self.manager.current = nome_tela


    def criar_botoes_para_matriculas(self, quantidade):
        self.botoes_matriculas.clear_widgets()
        self.matriculas.clear()

        for i in range(1, quantidade + 1):
            nome = f"Matrícula {i}"
            self.matriculas[nome] = {"campos": {}, "arquivos": {}}

            grupo = MDBoxLayout(orientation="horizontal", padding=10, spacing=10, size_hint_y=None)
            grupo.height = dp(240)

            campo_nome = MDTextField(
                hint_text=f"Nome da {nome}",
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            campo_valor = MDTextField(
                hint_text=f"Código da {nome}",
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            botao_planilha = MDButton(
                pos_hint={"center_x": 0.5},
                on_release=lambda x, n=nome: self.abrir_filemanager(n, "planilha")
            )
            botao_planilha.add_widget(MDButtonText(text="Selecionar Planilha"))

            botao_imagem = MDButton(
                pos_hint={"center_x": 0.5},
                on_release=lambda x, n=nome: self.abrir_filemanager(n, "imagem"),
            )
            botao_imagem.add_widget(MDButtonText(text="Selecionar Imagem"))

            botao_acao = MDButton(
                pos_hint={"center_x": 0.5},
                on_release=lambda x, n=nome: self.abrir_tela_detalhe(n)
            )
            botao_acao.add_widget(MDButtonText(text=f"Ação para {nome}"))

            self.matriculas[nome]["campos"] = {
                "nome": campo_nome,
                "valor": campo_valor
            }

            grupo.add_widget(campo_nome)
            grupo.add_widget(campo_valor)
            grupo.add_widget(botao_planilha)
            grupo.add_widget(botao_imagem)

            self.botoes_matriculas.add_widget(grupo)
            self.botoes_matriculas.add_widget(botao_acao)
