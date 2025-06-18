from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.screen import MDScreen
from app.screen3_dadosp.dados_function import extrair_dados_multiplos_pdfs, receber_dados_pdf

class ManagerScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        layout = BoxLayout(orientation="vertical")

        botao = MDButton(
            MDButtonText(text="Selecionar Arquivo"),
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
        self.file_manager.show("H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL")  

    def on_files_selected(self, paths):
        dados = extrair_dados_multiplos_pdfs(paths)
        tela_dados = self.manager.get_screen('dados')
        receber_dados_pdf(tela_dados, *dados)
        self.close_manager()
        self.manager.current = 'dados'  

    def close_manager(self, *args):
        self.file_manager.close()