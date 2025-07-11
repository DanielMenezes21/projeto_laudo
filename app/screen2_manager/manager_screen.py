from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.screen import MDScreen
from app.screen3_dadosp.dados_function import extrair_dados_multiplos_pdfs, receber_dados_pdf, extrair_solicitante_do_caminho
import os

class ManagerScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        layout = BoxLayout(orientation="vertical")

        botao = MDButton(
            MDButtonText(text="Selecionar Arquivo"),
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            on_release=lambda x: self.open_file_manager()
        )

        self.file_manager = MDFileManager(
            select_path=self.on_files_selected,
            exit_manager=self.close_manager,
            ext=[".pdf"],
            selector="multi"
        )

        layout.add_widget(botao)
        self.add_widget(layout)

    def open_file_manager(self):
        caminho_padrao = r"\\10.0.100.160\\Agropassos\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
        if os.path.exists(caminho_padrao):
            caminho = caminho_padrao
        else:
            caminho = os.path.expanduser("~/Documents")
        self.file_manager.show(caminho)

    def on_files_selected(self, paths):
        (nomes, cpfs, nomes_imoveis, municipio, estado, 
        latitudes, longitudes, dados_imoveis) = extrair_dados_multiplos_pdfs(paths)

        solicitante = extrair_solicitante_do_caminho(paths[0])
        print(f"👤 Solicitante extraído (manager): {solicitante}")
        
        tela_dados = self.manager.get_screen('dados')
        tela_dados.solicitante.text = solicitante
        tela_matricula = self.manager.get_screen('matricula')
        
        receber_dados_pdf(tela_dados, nomes, cpfs, nomes_imoveis, municipio, estado, dados_imoveis)
        tela_matricula.criar_botoes_para_matriculas(len(dados_imoveis))
        tela_matricula.receber_dados_imoveis(
            imoveis=nomes_imoveis,
            latitudes=latitudes,
            longitudes=longitudes,
            nomes_proprietarios=nomes, 
            dados_completos=dados_imoveis
        )
        
        self.close_manager()
        self.manager.current = 'dados'

    def close_manager(self, *args):
        self.file_manager.close()