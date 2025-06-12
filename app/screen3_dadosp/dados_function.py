import fitz
import re
import os
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run
from modules.data_atual import formatar_data
from kivymd.uix.dialog import MDDialog, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.widget import Widget
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.textfield import MDTextField

def go_next(self):
    # Atualiza o dicionário do proponente atual
    if self.proponente_atual in self.proponentes:
        self.proponentes[self.proponente_atual]["cpf"] = self.cpf.text
        self.proponentes[self.proponente_atual]["civil"] = self.civil.text
        self.proponentes[self.proponente_atual]["tratamento"] = self.tratamento

        proponente_data = self.proponentes[self.proponente_atual]
        campos = {
            "Tratamento": proponente_data.get("tratamento", ""),
            "Nome": self.proponente_atual,
            "CPF": proponente_data.get("cpf", ""),
            "Nome do Imóvel": self.nome_imovel.text,
            "Situação Civil": proponente_data.get("civil", ""),
            "Municipio": self.municipio.text,
            "Estado": self.estado.text,
            "Latitude": self.latitude.text,
            "Longitude": self.longitude.text,
            "Agencia": self.agencia.text
        }
    else:
        campos = {
            "Tratamento": self.tratamento,
            "Nome": self.proponente.text,
            "CPF": self.cpf.text,
            "Nome do Imóvel": self.nome_imovel.text,
            "Situação Civil": self.civil.text,
            "Municipio": self.municipio.text,
            "Estado": self.estado.text,
            "Latitude": self.latitude.text,
            "Longitude": self.longitude.text,
            "Agencia": self.agencia.text
        }

    """campos_vazios = [nome for nome, valor in campos.items() if not valor.strip()]

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
        return False"""

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
    tela_pdf.agencia = campos["Agencia"]

    self.manager.current = 'territorio'
    return True

def go_back(self):
    self.manager.current_screen.manager.current = "leitor"

def abrir_seletor_pdf(self, *args):
        data = formatar_data()
        data_nome = datetime.now()
        mes = f'{data_nome.month:02d}. {data.split("de")[1].strip()}'
        initial_path = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL"
        initial_path = os.path.join(initial_path, mes)
        self.file_manager.show(initial_path)

def fechar_arquivo(self, *args):
        self.file_manager.close()

def on_pdf_selecionado(self, caminho_pdf):
    fechar_arquivo(self)
    nomes, cpfs, nome_imovel, municipio, estado, latitude, longitude = extrair_dados_pdf(caminho_pdf)

    self.proponentes = {
        nome: {
            "cpf": cpf,
            "tratamento": "",
            "civil": ""
        }
        for nome, cpf in zip(nomes, cpfs)
    }

    self.menu_proponente.items = [
        {"text": nome, "on_release": lambda x=nome: self.selecionar_proponente(x)}
        for nome in self.proponentes
    ]

    self.nome_imovel.text = nome_imovel
    self.municipio.text = municipio
    self.estado.text = estado
    self.latitude.text = latitude
    self.longitude.text = longitude

    if self.menu_proponente.items:
        self.menu_proponente.open()

def extrair_dados_pdf(caminho_pdf):
    """
    Abre o PDF, varre todas as páginas em busca de Nome, CPF, Nome do Imóvel e Município.
    Retorna (nome, cpf, nome_imovel, municipio).
    """
    dados = {
        "nome": [],
        "cpf": [],
        "nome_imovel": "",
        "municipio": "",
        "estado": "",
        "latitude": "",
        "longitude": "",
    }

    padroes = {
        "nome": re.compile(r"\bNome:[:\-]?\s*(.+)", re.IGNORECASE),
        "cpf": re.compile(r"\bCPF[:\-]?\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})"),
        "nome_imovel": re.compile(r"\bNome do Imóvel Rural[:\-]?\s*(.+)", re.IGNORECASE),
        "municipio": re.compile(r"\bMunicípio[:\-]?\s*(.+)", re.IGNORECASE),
        "estado": re.compile(r"U[\r\n\u2028\u00a0]?F\s*[:\-]?\s*([^\r\n\u2028\u00a0]+)", re.IGNORECASE),
        "latitude": re.compile (r"\bLatitude:[:\-]?\s*(.+)", re.IGNORECASE),
        "longitude": re.compile (r"\bLongitude:[:\-]?\s(.+)", re.IGNORECASE)
    }

    try:
        pdf = fitz.open(caminho_pdf)
        for pagina in pdf:
            texto = pagina.get_text()
            texto = texto.replace('\u200b', '').replace('\xa0', ' ')
            nomes_encontrados = padroes["nome"].findall(texto)
            cpfs_encontrados = padroes["cpf"].findall(texto)
            dados["nome"].extend([n.strip() for n in nomes_encontrados if n.strip()])
            dados["cpf"].extend([c.strip() for c in cpfs_encontrados if c.strip()])
            for chave in ["nome_imovel", "municipio", "estado", "latitude", "longitude"]:
                if not dados[chave]:
                    m = padroes[chave].search(texto)
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
        dados["longitude"]
    )

def carregar_pdf_dados(self, caminho_car, caminho_cit):
    self.caminho_car = caminho_car
    self.caminho_cit = caminho_cit

    nome, cpf, nome_imovel, municipio, estado, latitude, longitude = extrair_dados_pdf(caminho_car)  
    self.proponente.text = nome
    self.cpf.text = cpf
    self.nome_imovel.text = nome_imovel
    self.municipio.text = municipio
    self.estado.text = estado
    self.latitude.text = latitude
    self.longitude.text = longitude

def abrir_dialogo_matriculas(self, *args):
    self.entrada_qtd_matriculas = MDTextField(
        hint_text="Quantidade de matrículas",
        input_filter="int",
        size_hint_x=0.9,
        pos_hint={"center_x": 0.5},
    )

    def confirmar_matriculas(x):
        qtd = self.entrada_qtd_matriculas.text
        if qtd.isdigit():
            qtd = int(qtd)
            self.manager.get_screen("matricula").criar_botoes_para_matriculas(qtd)
            self.dialogo_matriculas.dismiss()
            self.manager.current = "matricula"

    self.dialogo_matriculas = MDDialog(
        MDDialogHeadlineText(text="Número de Matrículas"),
        self.entrada_qtd_matriculas,
        MDDialogButtonContainer(
            MDButton(
                MDButtonText(text="OK"),
                on_release=confirmar_matriculas
            )
        )
    )
    self.dialogo_matriculas.open()

