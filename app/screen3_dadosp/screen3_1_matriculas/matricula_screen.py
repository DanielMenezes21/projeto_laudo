from kivymd.uix.screen import MDScreen
from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
from kivymd.uix.button import MDButton, MDButtonText, MDIconButton
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivy.uix.widget import Widget
from datetime import datetime
import openpyxl
import re
from modules.data_folder import formatar_data
import os
from kivy.metrics import dp
from app.screen3_dadosp.screen3_1_matriculas.screen3_1_1_detalhes.matricula_detalhe_screen import MatriculaDetalheScreen
from app.screen3_dadosp.screen3_1_matriculas.matricula_function import go_back

class MatriculaScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.layout = MDBoxLayout(orientation="vertical", padding=20, spacing=20)
        self.scroll = MDScrollView()
        self.botoes_matriculas = MDBoxLayout(orientation="vertical", spacing=10, size_hint_y=None)
        self.botoes_matriculas.bind(minimum_height=self.botoes_matriculas.setter("height"))
        self.scroll.add_widget(self.botoes_matriculas)

        self.button_back = MDIconButton(
            icon="arrow-left",
            size_hint=(1, None),
            pos_hint={"x": 0.06, "y": 0.4},
            size=(dp(56), dp(56)),
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            on_release=lambda x: go_back(self),
        )
        self.layout.add_widget(self.button_back)
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

    def extrair_numero_matricula_excel(self, caminho_arquivo):
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        aba = "AREA UTIL "
        if aba not in wb.sheetnames:
            print(f"{aba} não encontrado")
            return ""
        ws = wb[aba]
        pattern = re.compile(r"matr[ií]cula[\s\:\-\t]*([\d\.\,]+)", re.IGNORECASE)
        for row in ws.iter_rows(values_only=True):
            for idx, cell in enumerate(row):
                if isinstance(cell, str):
                    cell_limpa = cell.strip().replace('\n', '').replace('\r', '').replace('\t', ' ')
                    match = pattern.search(cell_limpa)
                    if match:
                        numero = match.group(1)
                        return numero
                    if "matr" in cell_limpa.lower() and idx + 1 < len(row):
                        prox = row[idx + 1]
                        if isinstance(prox, (int, float)):
                            numero = str(prox)
                            return numero
                if isinstance(cell, (int, float)) and idx > 0:
                    ant = row[idx - 1]
                    if isinstance(ant, str) and "matr" in ant.lower():
                        numero = str(cell)
                        return numero
        print("Número de matrícula não encontrado")
        return ""
    
    
    def receber_lat_long_pdf(self, latitude, longitude):
        """
        Preenche os campos de latitude e longitude das matrículas.
        """
        for i, nome_matricula in enumerate(self.matriculas):
            campos = self.matriculas[nome_matricula]["campos"]
            if isinstance(latitude, list) and i < len(latitude):
                campos["latitude"].text = str(latitude[i])
            else:
                campos["latitude"].text = str(latitude)
            if isinstance(longitude, list) and i < len(longitude):
                campos["longitude"].text = str(longitude[i])
            else:
                campos["longitude"].text = str(longitude)

    def extrair_valor_total_excel(self, caminho_arquivo):
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        aba = "SANEAMENTO"
        if aba not in wb.sheetnames:
            print(f"{aba} não encontrado")
            return ""
        ws = wb[aba]
        for row in ws.iter_rows(values_only=True):
            for idx, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip().lower() == "valor total":
                    for prox in row[idx+1:]:
                        if prox not in (None, "", "-"):
                            try:
                                valor = float(prox)
                                valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            except Exception:
                                valor_formatado = str(prox)
                            print(f"Valor Total encontrado: {valor_formatado}")
                            return valor_formatado
        print("Valor Total não encontrado")
        return ""
    
    def extrair_valor_liq_excel(self, caminho_arquivo):
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        aba = "LIQUIDAÇÃO"
        if aba not in wb.sheetnames:
            print(f"{aba} não encontrado")
            return ""
        ws = wb[aba]
        for row in ws.iter_rows(values_only=True):
            for idx, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip().lower() == "Valor de Liquidação Forçada":
                    for prox in row[idx+1:]:
                        if prox not in (None, "", "-"):
                            try:
                                valor = float(prox)
                                valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            except Exception:
                                valor_formatado = str(prox)
                            print(f"Valor de Liquidação Forçada encontrado: {valor_formatado}")
                            return valor_formatado
        print("Valor de Liquidação Forçada não encontrado")
        return ""

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

    def coletar_dados_matriculas(self):
        dados = []
        for nome_matricula, estrutura in self.matriculas.items():
            campos = estrutura["campos"]
            dados.append({
                "nome_imovel": campos["nome_imovel"].text,
                "matricula": campos["numero"].text,
                "valor_total": campos["valor"].text,
                "valor_liq": campos["valor_liq"].text,
                "latitude": campos["latitude"].text,
                "longitude": campos["longitude"].text
            })
        return dados

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

        if t == "planilha":
            numero = self.extrair_numero_matricula_excel(caminho)
            if numero:
                self.matriculas[m]["campos"]["numero"].text = numero
        valor_total = self.extrair_valor_total_excel(caminho)
        if valor_total:
            self.matriculas[m]["campos"]["valor"].text = str(valor_total)
        valor_liq = self.extrair_valor_liq_excel(caminho)
        if valor_liq:
            self.matriculas[m]["campos"]["valor_liq"].text = str(valor_liq)

    def abrir_tela_detalhe(self, nome_matricula):
        nome_tela = f"detalhe_{nome_matricula.replace(' ', '_').lower()}"
        if not self.manager.has_screen(nome_tela):
            nova_tela = MatriculaDetalheScreen(nome_matricula, name=nome_tela)
            self.manager.add_widget(nova_tela)
        dados = self.coletar_dados_matriculas()
        tela_pdf = self.manager.get_screen('pdf')
        self.manager.current = nome_tela

    def receber_dados_imoveis(self, imoveis, latitudes, longitudes, dados_completos=None):
        """Recebe os dados de imóveis de múltiplos PDFs"""
        self.dados_imoveis = dados_completos if dados_completos else {
            "imoveis": imoveis,
            "latitudes": latitudes,
            "longitudes": longitudes
        }
        
        print(f"📌 Total de imóveis recebidos: {len(imoveis)}")

        if hasattr(self, 'matriculas'):
            for i, nome_matricula in enumerate(self.matriculas):
                campos = self.matriculas[nome_matricula]["campos"]
                if i < len(imoveis):
                    campos["nome_imovel"].text = imoveis[i]
                if i < len(latitudes):
                    campos["latitude"].text = str(latitudes[i])
                if i < len(longitudes):
                    campos["longitude"].text = str(longitudes[i])

    def criar_botoes_para_matriculas(self, quantidade):
        self.botoes_matriculas.clear_widgets()
        self.matriculas.clear()

        for i in range(quantidade):
            nome = f"Matrícula {i+1}"
            self.matriculas[nome] = {"campos": {}, "arquivos": {}}

            nome_imovel = ""
            latitude = ""
            longitude = ""
            
            if hasattr(self, 'dados_imoveis'):
                if isinstance(self.dados_imoveis, list) and i < len(self.dados_imoveis):
                    nome_imovel = self.dados_imoveis[i].get('nome_imovel', "")
                    latitude = self.dados_imoveis[i].get('latitude', "")
                    longitude = self.dados_imoveis[i].get('longitude', "")
                elif isinstance(self.dados_imoveis, dict):
                    if i < len(self.dados_imoveis.get("imoveis", [])):
                        nome_imovel = self.dados_imoveis["imoveis"][i]
                    if i < len(self.dados_imoveis.get("latitudes", [])):
                        latitude = self.dados_imoveis["latitudes"][i]
                    if i < len(self.dados_imoveis.get("longitudes", [])):
                        longitude = self.dados_imoveis["longitudes"][i]

            grupo = MDBoxLayout(orientation="vertical", padding=10, spacing=25, size_hint_y=None)
            grupo.height = dp(650)  
            coords = MDBoxLayout(orientation="vertical", padding=0, spacing=5, size_hint_y=None)
            coords.height = 90

            campo_nome_imovel = MDTextField(
                MDTextFieldHintText(text=f"Nome do Imóvel {i}"),
                text=nome_imovel,
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            campo_numero = MDTextField(
                MDTextFieldHintText(text=f"Número da {nome}"),
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            campo_valor_total = MDTextField(
                MDTextFieldHintText(text=f"Valor total da {nome}"),
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            campo_valor_liq = MDTextField(
                MDTextFieldHintText(text=f"Valor líquido da {nome}"),
                size_hint=(0.9, None),
                height=50,
                pos_hint={"center_x": 0.5}
            )

            campo_latitude = MDTextField(
                MDTextFieldHintText(text=f"Latitude da {nome}"),
                text=latitude,
                size_hint=(0.9, None),
                height=40,
                pos_hint={"center_x": 0.5}
            )

            campo_longitude = MDTextField(
                MDTextFieldHintText(text=f"Longitude da {nome}"),
                text=longitude,
                size_hint=(0.9, None),
                height=40,
                pos_hint={"center_x": 0.5}
            )

            coords.add_widget(campo_latitude)
            coords.add_widget(campo_longitude)

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
                "nome_imovel": campo_nome_imovel,
                "numero": campo_numero,
                "valor": campo_valor_total,
                "valor_liq": campo_valor_liq,
                "latitude": campo_latitude,
                "longitude": campo_longitude
            }

            grupo.add_widget(campo_nome_imovel)
            grupo.add_widget(campo_numero)
            grupo.add_widget(campo_valor_total)
            grupo.add_widget(campo_valor_liq)
            grupo.add_widget(coords)
            grupo.add_widget(botao_planilha)
            grupo.add_widget(botao_imagem)

            self.botoes_matriculas.add_widget(grupo)
            self.botoes_matriculas.add_widget(botao_acao)