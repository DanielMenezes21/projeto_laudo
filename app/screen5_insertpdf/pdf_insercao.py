from kivy.metrics import dp
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from docx.shared import Cm
from docx.oxml.ns import qn
import os
import re
from app.screen5_insertpdf.pdf_extracao import extrair_paginas_como_imagens
from assets.document_style import *
from assets.document_page1 import *
from assets.document_page2 import *
from assets.document_page3 import *
from assets.document_page4 import *
from assets.document_page5 import *
from assets.document_page6 import *

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
                    novo_run.add_picture(imagem_path, width=Cm(14))
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

def inserir_imagem_no_placeholder(self, placeholder, caminho_imagem):
    for par in self.doc.paragraphs:
        if placeholder in par.text:
            for run in par.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, "")
            if os.path.exists(caminho_imagem):
                novo_run = par.add_run()
                novo_run.add_picture(caminho_imagem, width=Cm(14))
                print(f"✅ Imagem '{caminho_imagem}' inserida no placeholder '{placeholder}'")
            else:
                print(f"❌ Caminho inválido: {caminho_imagem}")
            return
    print(f"❌ Placeholder '{placeholder}' não encontrado no documento.")

def montar_documento(doc, qtd_imoveis):
    doc = configurar_documento()
    doc.add_page_break()

    for i in range(qtd_imoveis):
        doc = criar_titulo(doc)
        doc = adicionar_linha_fina(doc)
        doc = criar_secao_valor(doc)
        doc = adicionar_linha_fina(doc)
        doc = criar_secao_identificacao(doc)
        doc = adicionar_linha_fina(doc)
        doc = criar_secao_croqui(doc)
        doc = adicionar_linha_fina(doc)
        doc, tabela = geometria_terreno(doc)
        doc = adicionar_linha_fina(doc)
        doc = criar_secao_caracteristicas(doc)
        if i < qtd_imoveis:
            p = doc.add_paragraph()
            run = p.add_run(".")
            run.font.size = Pt(1)
            doc.add_page_break()
    for i in range(qtd_imoveis):
        doc = titulo(doc)
        doc = table_geo(doc)
        doc = adicionar_linha_fina(doc)
        doc = tabela_bioma(doc)
        doc = adicionar_linha_fina(doc)
        doc = area_APA(doc)
        doc = adicionar_linha_fina(doc)
        doc = table_passivo_ambiental(doc)
        if i ==qtd_imoveis -1:
            doc = campo_assinatura(doc)
            doc.add_page_break()
        if i < qtd_imoveis - 1:
            doc.add_page_break()
    doc = inserir_sumario(doc)
    doc.add_page_break()
    doc = adicionar_espaco(doc)
    doc = texto_solicitante(doc)
    doc = adicionar_espaco(doc)
    doc = texto_objetivo(doc)
    doc = adicionar_espaco(doc)
    doc = texto_finalidade(doc)
    doc = adicionar_espaco(doc)
    doc = texto_proprietario(doc)
    doc = adicionar_espaco(doc)
    doc = texto_ressalvas(doc)
    doc.add_page_break()
    doc = title_imovel(doc)
    doc = localizacao(doc)
    doc = acesso(doc)
    doc = carac_reg(doc)
    doc.add_page_break()
    doc = desc_imovel(doc)
    doc.add_page_break()

    return doc

def receber_dados_matriculas(self, lista_dados):
    self.lista_dados_matriculas = lista_dados
    self.qtd_imoveis = len(lista_dados)

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
        for i, dados in enumerate(self.lista_dados_matriculas):
            substituicoes = {
                "#TRATAMENTO": self.tratamento,
                "#PROPONENTE": self.nome,
                "#CPF_PROPONENTE": self.cpf,
                "#NOME_IMOVEL": dados.get("nome_imovel", ""),
                "#DATA_ATUAL": self.data_atual,
                "#CIVIL": self.civil,
                "#CIDADE_I": self.municipio,
                "#ESTADO_I": self.estado,
                "#LATITUDE": dados.get("latitude", ""),
                "#LONGITUDE": dados.get("longitude", ""),
                "#NMATRICULA": dados.get("matricula", ""),
                "#SOLICTANTE": self.solicitante,
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
        
        doc = configurar_documento()
        self.doc = montar_documento(doc, self.qtd_imoveis)

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