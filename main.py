from kivymd.app import MDApp
from kivy.uix.screenmanager import ScreenManager
from app.homepage.homepage import MenuScreen
from app.screen2_kml.leitor import LeitorScreen
from app.screen5_insertpdf.pdfscreen import PDFInsert
from app.screen3_dadosp.dadosscreen import DadosScreen
from kivy.config import Config

Config.set('kivy', 'exit_on_escape', '0')

class MainApp(MDApp):
    def build(self):
        self.title = "My KivyMD App"
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Blue"
        self.dados_extraidos = {}
        sm = ScreenManager()
        sm.add_widget(MenuScreen(name='main'))
        sm.add_widget(LeitorScreen(name='leitor'))
        sm.add_widget(PDFInsert(name='pdf'))
        sm.add_widget(DadosScreen(name='dados'))
        return sm

if __name__ == '__main__':
    MainApp().run()