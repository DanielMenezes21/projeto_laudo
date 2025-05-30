import fitz
import re
import os
from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run
from modules.data_atual import formatar_data
from kivymd.uix.dialog import MDDialog, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.widget import Widget
from kivymd.uix.button import MDButton, MDButtonText

def go_next(self):
    campos = {
        "Tratamento": self.tratamento,
        "Nome": self.proponente.text,
        "CPF": self.cpf.text,
        "Nome do Imóvel": self.nome_imovel.text,
        "Situação Civil": self.civil.text,
        "Municipio": self.municipio.text,
        "Estado": self.estado.text,
        "Latitude":self.latitude.text,
        "Longitude": self.longitude.text,
        "Matricula": self.matricula.text,
        "Agencia": self.agencia.text
    }

    campos_vazios = [nome for nome, valor in campos.items() if not valor.strip()]

    if campos_vazios:
        texto_erro = "Os seguintes campos estão vazios:\n" + "\n".join(f"- {campo}" for campo in campos_vazios)
        dialog = MDDialog(
            MDDialogHeadlineText(text="Campos obrigatórios não preenchidos"),
            MDDialogSupportingText(text=texto_erro),
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="OK"),
                    on_release=lambda x: dialog.dismiss()
                )
            )
        )
        dialog.open()
        return False

    tela_pdf = self.manager.get_screen('pdf')
    tela_pdf.tratamento = campos["Tratamento"]
    tela_pdf.nome = campos["Nome"]
    tela_pdf.cpf = campos["CPF"]
    tela_pdf.nome_imovel = campos["Nome do Imóvel"]
    tela_pdf.data_atual = formatar_data()
    tela_pdf.civil = campos["Situação Civil"]
    tela_pdf.municipio = campos["Municipio"]
    tela_pdf.estado = campos["Estado"]
    tela_pdf.latitude = campos["Latitude"]
    tela_pdf.longitude = campos["Longitude"]
    tela_pdf.matricula = campos["Matricula"]
    tela_pdf.agencia = campos["Agencia"]

    self.manager.current = 'territorio'
    return True

def go_back(self):
    self.manager.current_screen.manager.current = "leitor"

def abrir_seletor_pdf(self, *args):
        self.file_manager.show(os.getcwd())

def fechar_arquivo(self, *args):
        self.file_manager.close()

def on_pdf_selecionado(self, caminho_pdf):
        """
        Callback do FileManager. Recebe apenas 1 parâmetro,
        extrai nome/CPF do PDF selecionado e preenche os campos.
        """
        fechar_arquivo(self)
        nome, cpf, nome_imovel, municipio, estado, latitude, longitude, matricula = extrair_dados_pdf(caminho_pdf)
        self.proponente.text = nome
        self.cpf.text = cpf
        self.nome_imovel.text = nome_imovel
        self.municipio.text = municipio
        self.estado.text = estado
        self.latitude.text = latitude
        self.longitude.text = longitude
        self.matricula.text = matricula

def preencher_com_dados(self, caminho_car, caminho_cit):
        self.caminho_car = caminho_car
        self.caminho_cit = caminho_cit
        nome, cpf, nome_imovel, municipio, estado, latitude, longitude, matricula = extrair_dados_pdf(caminho_car)
        self.proponente.text = nome
        self.cpf.text = cpf
        self.nome_imovel.text = nome_imovel
        self.municipio.text = municipio
        self.estado.text = estado
        self.latitude.text = latitude
        self.longitude.text = longitude
        self.matricula.text = matricula

def extrair_dados_pdf(caminho_pdf):
    """
    Abre o PDF, varre todas as páginas em busca de Nome, CPF, Nome do Imóvel e Município.
    Retorna (nome, cpf, nome_imovel, municipio).
    """
    dados = {
        "nome": "",
        "cpf": "",
        "nome_imovel": "",
        "municipio": "",
        "estado": "",
        "latitude": "",
        "longitude": "",
        "matricula": ""
    }

    padroes = {
        "nome": re.compile(r"\bNome:[:\-]?\s*(.+)", re.IGNORECASE),
        "cpf": re.compile(r"\bCPF[:\-]?\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})"),
        "nome_imovel": re.compile(r"\bNome do Imóvel Rural[:\-]?\s*(.+)", re.IGNORECASE),
        "municipio": re.compile(r"\bMunicípio[:\-]?\s*(.+)", re.IGNORECASE),
        "estado": re.compile(r"U[\r\n\u2028\u00a0]?F\s*[:\-]?\s*([^\r\n\u2028\u00a0]+)", re.IGNORECASE),
        "latitude": re.compile (r"\bLatitude:[:\-]?\s*(.+)", re.IGNORECASE),
        "longitude": re.compile (r"\bLongitude:[:\-]?\s(.+)", re.IGNORECASE),
        "matricula": re.compile(r"(?i)Município do Cartório[\r\n\u2028\u00a0]+([^\r\n]+)", re.IGNORECASE)
    }

    try:
        pdf = fitz.open(caminho_pdf)
        for pagina in pdf:
            texto = pagina.get_text()
            texto = texto.replace('\u200b', '').replace('\xa0', ' ')
            for chave, regex in padroes.items():
                if not dados[chave]:
                    m = regex.search(texto)
                    if m:
                        dados[chave] = m.group(1).strip()
            if all(dados.values()):
                break
        pdf.close()
    except Exception as e:
        print(f"❌ Erro ao extrair dados do PDF: {e}")

    return (
        dados["nome"],
        dados["cpf"],
        dados["nome_imovel"],
        dados["municipio"],
        dados["estado"],
        dados["latitude"],
        dados["longitude"],
        dados["matricula"]
    )

def carregar_pdf_dados(self, caminho_car, caminho_cit):
    self.caminho_car = caminho_car
    self.caminho_cit = caminho_cit

    nome, cpf, nome_imovel, municipio, estado, latitude, longitude, matricula = extrair_dados_pdf(caminho_car)  
    self.proponente.text = nome
    self.cpf.text = cpf
    self.nome_imovel.text = nome_imovel
    self.municipio.text = municipio
    self.estado.text = estado
    self.latitude.text = latitude
    self.longitude.text = longitude
    self.matricula.text = matricula
