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
        return

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

    self.manager.current = 'pdf'

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
        "matricula": re.compile(r"(?i)Número da Matrícula[ \n]+Data do Documento[ \n]+Livro[ \n]+Folha[ \n]+Município do Cartório[\r\n\u2028\u00a0]+([^\r\n]+)", re.IGNORECASE)
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

    placeholder_car = "#SUBSTITUIR_CAR"  
    self.inserir_pdf_no_word(self.caminho_car, placeholder_car)

def receber_dados(self):
    tratamento = self.tratamento
    nome = self.proponente.text
    cpf = self.cpf.text
    nome_imovel = self.nome_imovel.text
    data_atual = formatar_data()
    civil = self.civil.text
    municipio = self.municipio.text
    estado = self.estado.text
    latitude = self.latitude.text
    longitude = self.longitude.text
    matricula = self.matricula.text
    agencia = self.agencia.text

    modelo_path = os.path.join(os.getcwd(), "models", "MODELO_LAUDO.docx")
    doc = Document(modelo_path)

    substituicoes = {
        "#TRATAMENTO": tratamento,
        "#PROPONENTE": nome,
        "#CPF_PROPONENTE": cpf,
        "#NOME_IMOVEL": nome_imovel,
        "#DATA_ATUAL": data_atual,
        "#CIVIL": civil,
        "#CIDADE_I": municipio,
        "#ESTADO_I": estado,
        "#LATITUDE": latitude,
        "#LONGITUDE": longitude,
        "#NMATRICULA": matricula,
        "#AGENCIA": agencia
    }

    def substituir_em_runs(par):
        texto_completo = ''.join(run.text for run in par.runs)
        novo_texto = texto_completo
        for chave, valor in substituicoes.items():
            novo_texto = novo_texto.replace(chave, valor)

        if novo_texto != texto_completo:
            for run in par.runs[:]:
                run._element.getparent().remove(run._element)
            par.add_run(novo_texto)

    def substituir_em_paragrafos(paragrafos):
        for par in paragrafos:
            texto_completo = ''.join(run.text for run in par.runs)
            novo_texto = texto_completo

            for chave, valor in substituicoes.items():
                if chave in novo_texto:
                    novo_texto = novo_texto.replace(chave, valor)

            if novo_texto != texto_completo:
                for run in par.runs[:]:
                    run._element.getparent().remove(run._element)
                par.add_run(novo_texto)
            if chave in texto_completo:
                print(f"Substituindo '{chave}' por '{valor}' em: {texto_completo}")


    def substituir_em_tabela(tabela):
        for linha in tabela.rows:
            for celula in linha.cells:
                substituir_em_paragrafos(celula.paragraphs)
                for tabela_interna in celula.tables:
                    substituir_em_tabela(tabela_interna)

    substituir_em_paragrafos(doc.paragraphs)

    for tabela in doc.tables:
        substituir_em_tabela(tabela)

    for section in doc.sections:
        substituir_em_paragrafos(section.header.paragraphs)
        substituir_em_paragrafos(section.footer.paragraphs)

    for shape in doc.inline_shapes:
        if shape._inline.graphic.graphicData.uri.endswith("/wordprocessingShape"):
            for box in shape._inline.graphic.graphicData.xpath(".//w:txbxContent"):
                for par_el in box.iter(qn('w:p')):
                    for r in par_el.iter(qn('w:t')):
                        if r.text:
                            for chave, valor in substituicoes.items():
                                if chave in r.text:
                                    r.text = r.text.replace(chave, valor)