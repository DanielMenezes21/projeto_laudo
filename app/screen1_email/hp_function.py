from kivy.clock import Clock
from threading import Thread
from modules.automacao import EmailAutomator
from kivy.metrics import dp
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivymd.uix.dialog import MDDialog, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.button import MDButton, MDButtonText
def next_screen(self, instance):
        self.manager.current_screen.manager.current = "leitor"

def download_file(self, instance):
        self.progress.active = True

        gmail_service = EmailAutomator.login_gmail()
        query = self.text_field.text
        if not query:
            query = " " 

        def run_download():
            EmailAutomator.baixar_anexos(gmail_service, query=query)
            Clock.schedule_once(lambda dt: after_download(self), 5)

        Thread(target=run_download).start()

def after_download(self):
        self.progress.active = False
        show_file_manager(self)  

def show_file_manager(self):
        initial_path = r"C:\Users\DESKTOP\Desktop\automacao_laudo\anexos"
        self.file_manager = MDFileManager(
            exit_manager=lambda *args: close_file_manager(self, *args),
            select_path=lambda path: select_pdf_file(self, path),
        )
        self.file_manager.show(initial_path)
        Window.bind(on_keyboard=handle_keyboard)

def handle_keyboard(self, instance, keyboard, keycode, text, modifiers):
        if keyboard == 27 and self.file_manager:
            close_file_manager(self)
            return True
        return False

def close_file_manager(self, *args):
        if self.file_manager:
            self.file_manager.close()
            self.file_manager = None
        Window.unbind(on_keyboard=handle_keyboard)

def select_pdf_file(self, path):
        from modules.leitorpdf import extrair_coordenadas_pdf, gerar_kml
        coords = extrair_coordenadas_pdf(path)
        if coords:
            gerar_kml(coords, path)
        else:
            dialog = MDDialog(
                MDDialogHeadlineText(text="Erro"),
                MDDialogSupportingText(text="Não foi possível extrair as coordenadas do PDF"),
                MDDialogButtonContainer(
                    MDButton(
                        MDButtonText(text="OK"),
                        on_release=lambda x: dialog.dismiss()
                    )
                )
            )
            dialog.open()
            return 
        close_file_manager(self)