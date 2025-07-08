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
    if self.proponente_atual in self.proponentes:
        self.proponentes[self.proponente_atual]["cpf"] = self.cpf.text
        self.proponentes[self.proponente_atual]["civil"] = self.civil.text
        self.proponentes[self.proponente_atual]["tratamento"] = self.tratamento

        proponente_data = self.proponentes[self.proponente_atual]
        campos = {
            "Tratamento": proponente_data.get("tratamento", ""),
            "Nome": self.proponente_atual,
            "CPF": proponente_data.get("cpf", ""),
            "Situação Civil": proponente_data.get("civil", ""),
            "Municipio": self.municipio.text,
            "Estado": self.estado.text,
            "Solicitante": self.solicitante.text
        }
    else:
        campos = {
            "Tratamento": self.tratamento,
            "Nome": self.proponente.text,
            "CPF": self.cpf.text,
            "Situação Civil": self.civil.text,
            "Municipio": self.municipio.text,
            "Estado": self.estado.text,
            "Solicitante": self.solicitante.text
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
    print("🔍 Dados enviados à tela PDF:")
    for chave, valor in campos.items():
        print(f"{chave}: {valor}")
    tela_pdf = self.manager.get_screen('pdf')
    tela_pdf.tratamento = campos["Tratamento"]
    tela_pdf.nome = campos["Nome"]
    tela_pdf.cpf = campos["CPF"]
    tela_pdf.data_atual = formatar_data()
    tela_pdf.civil = campos["Situação Civil"]
    tela_pdf.municipio = campos["Municipio"]
    tela_pdf.estado = campos["Estado"]
    tela_pdf.solicitante = campos["Solicitante"]

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
        if not os.path.exists(initial_path):
            initial_path = os.path.join(os.path.expanduser("~/Documents"), mes)
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
            "civil": "",
            "nome_imovel": nome_imovel
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

def extrair_dados_multiplos_pdfs(lista_caminhos_pdf):
    dados_imoveis = []
    municipio = ""
    estado = ""

    for caminho_pdf in lista_caminhos_pdf:
        nomes, cpfs, nome_imovel, municipio_, estado_, latitude, longitude = extrair_dados_pdf(caminho_pdf)
        
        dados_imoveis.append({
            'nomes': nomes,
            'cpfs': cpfs,
            'nome_imovel': nome_imovel if nome_imovel else f"Imóvel {len(dados_imoveis)+1}",
            'latitude': latitude[0] if latitude else "",
            'longitude': longitude[0] if longitude else "",
            'proprietario': ", ".join(nomes) if nomes else "" ,
        })

        if not municipio and municipio_:
            municipio = municipio_
        if not estado and estado_:
            estado = estado_

    todos_nomes = [nome for imovel in dados_imoveis for nome in imovel['nomes']]
    todos_cpfs = [cpf for imovel in dados_imoveis for cpf in imovel['cpfs']]
    nomes_imoveis = [imovel['nome_imovel'] for imovel in dados_imoveis]
    latitudes = [imovel['latitude'] for imovel in dados_imoveis]
    longitudes = [imovel['longitude'] for imovel in dados_imoveis]

    return (todos_nomes, todos_cpfs, nomes_imoveis, municipio, estado, latitudes, longitudes, dados_imoveis)

def extrair_dados_pdf(caminho_pdf):
    """
    Abre o PDF, varre todas as páginas em busca de Nome, CPF, Nome do Imóvel, Município, Latitude e Longitude.
    Retorna (nome, cpf, nome_imovel, municipio, estado, latitude, longitude).
    """
    dados = {
        "nome": [],
        "cpf": [],
        "nome_imovel": "",
        "municipio": "",
        "estado": "",
        "latitude": [],
        "longitude": [],
    }

    padroes = {
        "nome": re.compile(r"\bNome:[:\-]?\s*(.+)", re.IGNORECASE),
        "cpf": re.compile(r"\b(CPF|CNPJ)[:\-]?\s*((?:\d{3}[.\s]?){2}\d{3}[-\s]?\d{2}|\d{2}[.\s]?\d{3}[.\s]?\d{3}[\/\s]?\d{4}[-\s]?\d{2})", re.IGNORECASE),
        "nome_imovel": re.compile(r"\bNome do Imóvel Rural[:\-]?\s*(.+)", re.IGNORECASE),
        "municipio": re.compile(r"\bMunicípio[:\-]?\s*(.+)", re.IGNORECASE),
        "estado": re.compile(r"U[\r\n\u2028\u00a0]?F\s*[:\-]?\s*([^\r\n\u2028\u00a0]+)", re.IGNORECASE),
        "latitude": re.compile(r"\bLatitude:[:\-]?\s*(.+)", re.IGNORECASE),
        "longitude": re.compile(r"\bLongitude:[:\-]?\s(.+)", re.IGNORECASE)
    }

    try:
        pdf = fitz.open(caminho_pdf)
        for pagina in pdf:
            texto = pagina.get_text()
            texto = texto.replace('\u200b', '').replace('\xa0', ' ')
            
            nomes_encontrados = padroes["nome"].findall(texto)
            cpfs_encontrados = padroes["cpf"].findall(texto)
            dados["nome"].extend([n.strip() for n in nomes_encontrados if n.strip()])
            dados["cpf"].extend([c[1].strip() for c in cpfs_encontrados if c[1].strip()])
            
            latitudes_encontradas = padroes["latitude"].findall(texto)
            longitudes_encontradas = padroes["longitude"].findall(texto)
            dados["latitude"].extend([l.strip() for l in latitudes_encontradas if l.strip()])
            dados["longitude"].extend([l.strip() for l in longitudes_encontradas if l.strip()])
            
            if latitudes_encontradas or longitudes_encontradas:
                print(f"\n📌 Coordenadas encontradas (Página {pagina.number + 1}):")
                for lat, lon in zip(latitudes_encontradas, longitudes_encontradas):
                    print(f"  → Latitude: {lat.strip()}")
                    print(f"  → Longitude: {lon.strip()}")
            
            for chave in ["nome_imovel", "municipio", "estado"]:
                if not dados[chave]:
                    m = padroes[chave].search(texto)
                    if m:
                        dados[chave] = m.group(1).strip()
            
            if all(dados.values()):
                break
        
        pdf.close()
        
        print("\n✅ Dados consolidados do PDF:")
        print(f"  - Nomes: {dados['nome']}")
        print(f"  - CPFs: {dados['cpf']}")
        print(f"  - Nome do Imóvel: {dados['nome_imovel']}")
        print(f"  - Município: {dados['municipio']}")
        print(f"  - Estado: {dados['estado']}")
        print(f"  - Latitudes: {dados['latitude']}")
        print(f"  - Longitudes: {dados['longitude']}")
        
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

def receber_dados_pdf(tela_dados, nomes, cpfs, nomes_imovel, municipio, estado, dados_imoveis=None):
    tela_dados.proponentes = {
        nome: {
            "nome": nome,
            "cpf": cpf,
            "tratamento": "",
            "civil": "",
            "nome_imovel": nomes_imovel[i] if i < len(nomes_imovel) else ""
        }
        for i, (nome, cpf) in enumerate(zip(nomes, cpfs))
    }

    tela_dados.menu_proponente.items = [
        {"text": nome, "on_release": lambda x=nome: tela_dados.selecionar_proprietario(x)}
        for nome in tela_dados.proponentes
    ]

    tela_dados.municipio.text = municipio
    tela_dados.estado.text = estado

    if tela_dados.menu_proponente.items:
        tela_dados.menu_proponente.open()
    
    if dados_imoveis:
        tela_dados.dados_imoveis = dados_imoveis

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

    self.proponentes = {
    nome[0] if isinstance(nome, list) else nome: {
        "cpf": cpf[0] if isinstance(cpf, list) else cpf,
        "nome_imovel": nome_imovel
    }
}

def abrir_dialogo_matriculas(self, *args):
    if hasattr(self, 'dados_imoveis'):  
        qtd_imoveis = len(self.dados_imoveis)
        print(f"📄 Número de propriedades detectadas: {qtd_imoveis}")
        tela_matricula = self.manager.get_screen('matricula')
        tela_final = self.manager.get_screen('pdf')
        tela_final.qtd_imoveis = qtd_imoveis
        tela_final.dados_imoveis = self.dados_imoveis
        tela_matricula.receber_dados_imoveis(
            imoveis=[imovel['nome_imovel'] for imovel in self.dados_imoveis],
            latitudes=[imovel['latitude'] for imovel in self.dados_imoveis],
            longitudes=[imovel['longitude'] for imovel in self.dados_imoveis],
            nomes_proprietarios=[imovel.get('proprietario', '') for imovel in self.dados_imoveis],
            dados_completos=self.dados_imoveis  
        )

        tela_matricula.criar_botoes_para_matriculas(qtd_imoveis)
        self.manager.current = "matricula"
    else:
        self.entrada_qtd_matriculas = MDTextField(
            hint_text="Quantidade de matrículas",
            input_filter="int",
            size_hint_x=0.9,
            pos_hint={"center_x": 0.5},
        )

        def confirmar_matriculas(x):
            if self.entrada_qtd_matriculas.text.isdigit():
                qtd = int(self.entrada_qtd_matriculas.text)
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
