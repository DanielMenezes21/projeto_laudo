from kivymd.app import MDApp
from kivy.uix.screenmanager import ScreenManager
from app.screen1_email.homepage import MenuScreen
from app.screen2_kml.leitor import LeitorScreen
from app.screen2_manager.manager_screen import ManagerScreen
from app.screen3_dadosp.dadosscreen import DadosScreen
from app.screen3_dadosp.screen3_1_matriculas.matricula_screen import MatriculaScreen
from app.screen4_territorio.soloscreen import SoloScreen
from app.screen5_insertpdf.pdfscreen import PDFInsert
from kivy.uix.screenmanager import (SlideTransition, 
FadeTransition, SwapTransition, WipeTransition,
FallOutTransition, RiseInTransition)
from kivy.core.window import Window
from modules.validacao import VALIDACAO_ESPECIFICA
import os
import sys

from kivy.config import Config

Config.set('kivy', 'exit_on_escape', '0')
Config.set('input', 'mouse', 'mouse,multitouch_on_demand')

class MainApp(MDApp):
    def build(self):
        self.title = "My KivyMD App"
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Blue"
        self.dados_extraidos = {}
        self.lista_dados_matriculas = []
        sm = ScreenManager(transition=FadeTransition())
        sm.add_widget(MenuScreen(name='main'))
        sm.add_widget(LeitorScreen(name='leitor'))
        sm.add_widget(MatriculaScreen(name="matricula", lista_dados_matriculas=self.lista_dados_matriculas))
        sm.add_widget(ManagerScreen(name="manager"))
        sm.add_widget(DadosScreen(name='dados'))
        sm.add_widget(SoloScreen(name='territorio'))
        sm.add_widget(PDFInsert(name='pdf'))
        return sm

if __name__ == '__main__':
    MainApp().run()