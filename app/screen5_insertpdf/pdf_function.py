import os
import re
from docx.shared import Inches
from docx.oxml.ns import qn
import uuid
import fitz
from docx import Document
from docx.shared import Inches
from pdf2image import convert_from_path
from kivymd.uix.list import MDListItem, MDListItemHeadlineText
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.metrics import dp
from kivy.uix.popup import Popup
from app.screen5_insertpdf.pdf_function2 import extrair_paginas_como_imagens
from app.screen3_dadosp.dados_function import *

def go_back(self, *args):
    self.manager.current_screen.manager.current = "territorio"

def selecionar_pdf_car(self, caminho_pdf):
    self.caminho_car = caminho_pdf
    adicionar_item(self, caminho_pdf, "car")  

def selecionar_pdf_cit(self, caminho_pdf):
    self.caminho_cit = caminho_pdf
    adicionar_item(self, caminho_pdf, "cit")  

def carregar_estrutura_pasta(self, caminho_pasta, tipo):
    """Função para carregar arquivos PDF dentro de uma pasta"""
    for nome in sorted(os.listdir(caminho_pasta)):
        caminho_completo = os.path.join(caminho_pasta, nome)
        if os.path.isdir(caminho_completo):
            match = re.match(r"PROCESSO\s*Nº\s*(\d+)", nome)
            if match:
                numero_processo = match.group(1)
                self.numero_processo = numero_processo  
                if numero_processo == "":
                    MDSnackbar(
                        MDSnackbarText(text="Número do processo não encontrado!"),
                        y=dp(24)
                    ).open()
                    return
                carregar_estrutura_pasta(self, caminho_completo, tipo)
        elif os.path.isfile(caminho_completo) and nome.lower().endswith(".pdf"):
            adicionar_item(self, caminho_completo, tipo)

def adicionar_item(self, caminho_pdf, tipo):
    texto = f"Arquivo: {os.path.basename(caminho_pdf)}"
    if tipo == "car":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_car(self, f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CAR] {texto}"))
        self.file_list.add_widget(item)
    elif tipo == "cit":
        item = MDListItem(
            on_release=lambda x, f=caminho_pdf: selecionar_pdf_cit(self,f)
        )
        item.add_widget(MDListItemHeadlineText(text=f"[CIT] {texto}"))
        self.file_list2.add_widget(item)

def carregar_estrutura(self, tipo):
    print(f"Carregando estrutura para: {tipo}")
    if not os.path.exists(self.current_path):
        MDSnackbar(
            MDSnackbarText(text="Caminho não encontrado!"),
            y=dp(24)
        ).open()
        return

    if tipo == "car":
        self.file_list.clear_widgets()
    elif tipo == "cit":
        self.file_list2.clear_widgets()

    print(f"Listando arquivos no diretório: {self.current_path}")
    arquivos_adicionados = 0

    for nome in sorted(os.listdir(self.current_path)):
        caminho_completo = os.path.join(self.current_path, nome)

        if os.path.isdir(caminho_completo):
            carregar_estrutura_pasta(self,caminho_completo, tipo)
        elif os.path.isfile(caminho_completo) and nome.lower().endswith(".pdf"):
            adicionar_item(self,caminho_completo, tipo)
            arquivos_adicionados += 1

    if arquivos_adicionados == 0:
        MDSnackbar(
            MDSnackbarText(text="Nenhum arquivo PDF encontrado na pasta."),
            y=dp(24)
        ).open()


def load_directory(self,tipo, *args):
    print(f"Carregando diretório: {self.current_path}")
    self.pastas_abertas = set() 
    carregar_estrutura(self, tipo)

def entrar_em_pasta(self, caminho_pasta, tipo):
    if caminho_pasta in self.pastas_abertas:
        self.pastas_abertas.remove(caminho_pasta)
    else:
        self.pastas_abertas.add(caminho_pasta)
    carregar_estrutura(self, tipo)

def go_up(self, *args):
    self.current_path = os.path.dirname(self.current_path)
    if not os.path.exists(self.current_path):
        MDSnackbar(
            MDSnackbarText(text="Caminho não encontrado!"),
            y=dp(24)
        ).open()
        return
    carregar_estrutura(self)

def initialize_word(self):
    modelo_path = os.path.join(os.getcwd(), "models","MODELO_LAUDO.docx")
    if not os.path.exists(modelo_path):
        MDSnackbar(
            MDSnackbarText(text="Modelo de documento não encontrado!"),
            y=dp(24)
        ).open()
        return
    self.doc = Document(modelo_path)

def inserir_pdf_no_word(self, caminho_pdf, placeholder):
    if not hasattr(self, "doc"):
        print("Documento Word não inicializado!")
        return
    else:
        print("Documento iniciado")

    def substituir_em_paragrafos(paragrafos):
        for par in paragrafos:
            texto_completo = ''.join(run.text for run in par.runs)
            if placeholder in texto_completo:
                for run in par.runs:
                    run.text = ""
                for imagem_path in imagens:
                    novo_run = par.add_run()
                    novo_run.add_picture(imagem_path, width=Inches(6))
                return True
        return False

    try:
        imagens = extrair_paginas_como_imagens(caminho_pdf)
        if not imagens:
            MDSnackbar(
                MDSnackbarText(text="Nenhuma imagem encontrada no PDF."),
                y=dp(24)
            ).open()
            return

        encontrado = substituir_em_paragrafos(self.doc.paragraphs)

        if not encontrado:
            for tabela in self.doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        if substituir_em_paragrafos(celula.paragraphs):
                            encontrado = True
                            break
                    if encontrado:
                        break
                if encontrado:
                    break

        if encontrado:
            if not hasattr(self, "pdfs_inseridos"):
                self.pdfs_inseridos = set()
            self.pdfs_inseridos.add(placeholder)

    except Exception as e:
        MDSnackbar(
            MDSnackbarText(text=f"Erro ao inserir PDF: {str(e)}"),
            y=dp(24)
        ).open()

def substituir_texto_formatado(paragrafos, substituicoes):
    for par in paragrafos:
        for run in par.runs:
            for placeholder, valor in substituicoes.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, valor)

def inserir_imagem_no_placeholder(self, placeholder, caminho_imagem):
    for par in self.doc.paragraphs:
        if placeholder in par.text:
            for run in par.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, "")
            if os.path.exists(caminho_imagem):
                novo_run = par.add_run()
                novo_run.add_picture(caminho_imagem, width=Inches(6))
                print(f"✅ Imagem '{caminho_imagem}' inserida no placeholder '{placeholder}'")
            else:
                print(f"❌ Caminho inválido: {caminho_imagem}")
            return
    print(f"❌ Placeholder '{placeholder}' não encontrado no documento.")

def gerar_documento(self):
    try:
        processo = ""
        if self.current_path:
            pasta_anexos = self.current_path  
            if os.path.exists(pasta_anexos):
                for subpasta in os.listdir(pasta_anexos):
                    subpasta_completa = os.path.join(pasta_anexos, subpasta)
                    if os.path.isdir(subpasta_completa):
                        match = re.search(r"(?i)processo\s*n[°º]\s*(\d+)", subpasta, re.IGNORECASE)
                        if match:
                            processo = match.group(1)

        substituicoes = {
            "#TRATAMENTO": self.tratamento,
            "#PROPONENTE": self.nome,
            "#CPF_PROPONENTE": self.cpf,
            "#NOME_IMOVEL": self.nome_imovel,
            "#DATA_ATUAL": self.data_atual,
            "#CIVIL": self.civil,
            "#CIDADE_I": self.municipio,
            "#ESTADO_I": self.estado,
            "#LATITUDE": self.latitude,
            "#LONGITUDE": self.longitude,
            "#NMATRICULA": self.matricula,
            "#AGENCIA": self.agencia,
            "#DESCRICAO_IMOVEL": self.descricao_imovel,
            "#REGIAO_CIDADE": self.descricao_cidade,
            "#ATIVIDADE_IMOVEL": self.atividade_imovel,
            "#REGIAO_IMOVEL": self.regiao_imovel,
            "#DECLIVIDADE_I": self.declividade,
            "#HIDROGRAFIA_I": self.hidrografia,
            "#TIPO_SOLO": self.resumo_solo,
            "#DESCRICAO_SOLO": self.texto_solos,
            "#ROTA_ACESSO": self.rotas,
            "#NPROCESSO": processo,
        }

        for chave, valor in substituicoes.items():
            print(f"{chave}: {type(valor)}")

        if hasattr(self, "caminho_declividade"):
            inserir_imagem_no_placeholder(self, "#IMAGEM_DECLIVIDADE", self.caminho_declividade)
        if hasattr(self, "caminho_hidrografia"):
            inserir_imagem_no_placeholder(self, "#IMAGEM_HIDROGRAFIA", self.caminho_hidrografia)
        if hasattr(self, "caminho_rotas"):
            inserir_imagem_no_placeholder(self, "#IMAGEM_ACESSO", self.caminho_rotas)

        substituicoes = { chave: (valor if isinstance(valor, str) else str(valor))
            for chave, valor in substituicoes.items() }
        
        def substituir_em_runs(par):
            for run in par.runs:
                for chave, valor in substituicoes.items():
                    if chave in run.text:
                        run.text = run.text.replace(chave, valor)

        def substituir_em_paragrafos(paragrafos):
            for par in paragrafos:
                substituir_em_runs(par)

        def substituir_em_tabela(tabela):
            for linha in tabela.rows:
                for celula in linha.cells:
                    substituir_em_paragrafos(celula.paragraphs)
                    for tabela_interna in celula.tables:
                        substituir_em_tabela(tabela_interna)

        substituir_em_paragrafos(self.doc.paragraphs)

        for tabela in self.doc.tables:
            substituir_em_tabela(tabela)

        for section in self.doc.sections:
            substituir_em_paragrafos(section.header.paragraphs)
            substituir_em_paragrafos(section.footer.paragraphs)

        for shape in self.doc.inline_shapes:
            if shape._inline.graphic.graphicData.uri.endswith("/wordprocessingShape"):
                for box in shape._inline.graphic.graphicData.xpath(".//w:txbxContent"):
                    for par_el in box.iter(qn('w:p')):
                        for r in par_el.iter(qn('w:t')):
                            if r.text:
                                for chave, valor in substituicoes.items():
                                    if chave in r.text:
                                        r.text = r.text.replace(chave, valor)

        if self.caminho_car:
            inserir_pdf_no_word(self, self.caminho_car, "#SUBSTITUIR_CAR")
        if self.caminho_cit:
            inserir_pdf_no_word(self, self.caminho_cit, "#SUBSTITUIR_CIT")

        nome_arquivo = f"LAUDO DE AVALIAÇÃO {processo} {self.nome}.docx"
        output_path = os.path.join(os.getcwd(), "output", nome_arquivo)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        self.doc.save(output_path)
        os.startfile(output_path)

        MDSnackbar(
            MDSnackbarText(text="✅Documento gerado com sucesso!"),
            y=dp(24)
        ).open()

    except Exception as e:
        print(f"❌ Erro ao gerar documento: {e}")
        MDSnackbar(
            MDSnackbarText(text=f"Erro: {str(e)}"),
            y=dp(24)
        ).open()
