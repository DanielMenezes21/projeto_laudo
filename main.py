from kivymd.app import MDApp
from kivy.uix.screenmanager import ScreenManager
from app.screen1_email.homepage import MenuScreen
from app.screen2_kml.leitor import LeitorScreen
from app.screen3_dadosp.dadosscreen import DadosScreen
from app.screen4_territorio.soloscreen import SoloScreen
from app.screen5_insertpdf.pdfscreen import PDFInsert
from kivy.uix.screenmanager import (SlideTransition, 
FadeTransition, SwapTransition, WipeTransition,
FallOutTransition, RiseInTransition)
from kivy.core.window import Window
from modules.validacao import VALIDACAO_ESPECIFICA

from kivy.config import Config

Config.set('kivy', 'exit_on_escape', '0')
class Gerenciador(ScreenManager):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.telas = ['main', 'leitor', 'dados', 'territorio', 'pdf']
        self.current_index = 0
        Window.bind(on_key_down=self.on_key_down)

    def on_key_down(self, window, key, *args):
        tela_atual = self.get_screen(self.current)

        if key == 275:  
            nome_tela = self.current
            if nome_tela in VALIDACAO_ESPECIFICA:
                if not VALIDACAO_ESPECIFICA[nome_tela](tela_atual):
                    return True  
            if self.current_index < len(self.telas) - 1:
                self.current_index += 1
                self.current = self.telas[self.current_index]
            return True

        elif key == 276:
            if self.current_index > 0:
                self.current_index -= 1
                self.current = self.telas[self.current_index]
            return True

        return False
        
class MainApp(MDApp):
    def build(self):
        self.title = "My KivyMD App"
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Blue"
        self.dados_extraidos = {}
        sm = Gerenciador(transition=FadeTransition())
        sm.add_widget(MenuScreen(name='main'))
        sm.add_widget(LeitorScreen(name='leitor'))
        sm.add_widget(DadosScreen(name='dados'))
        sm.add_widget(SoloScreen(name='territorio'))
        sm.add_widget(PDFInsert(name='pdf'))
        return sm

if __name__ == '__main__':
    MainApp().run()