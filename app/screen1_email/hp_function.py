from kivy.clock import Clock
from threading import Thread
from modules.automacao import EmailAutomator
from kivy.metrics import dp
from kivymd.uix.filemanager import MDFileManager
from kivy.core.window import Window
from kivymd.uix.dialog import MDDialog, MDDialogSupportingText, MDDialogHeadlineText, MDDialogButtonContainer
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.selectioncontrol import MDCheckbox
from modules.data_folder import formatar_data
from datetime import datetime
import os
from app.screen1_email.checkbox import add_checkboxes_to_pdfs
from kivy.clock import mainthread

data = formatar_data()
data_nome = datetime.now()
mes = f'{data_nome.month:02d}. {data.split('de')[1].strip()}'

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
    self.selected_files = []  
    initial_path = r"H:\1. AVALIAÇÕES\01. AVALIAÇÕES SICREDI\01. RURAL"
    initial_path = os.path.join(initial_path, mes)
    self.file_manager = MDFileManager(
        exit_manager=lambda *args: close_file_manager(self, *args),
        select_path=lambda path: select_pdf_file(self, path),
        preview=False,
        icon_selection_button="check",
        selection_button=True,
        selector="multi",
        background_color_selection_button=(0, 0, 0, 0),
        icon_color=(1, 1, 1, 1),
        background_color_toolbar="brown"
    )
    self.file_manager.show(initial_path)
    Window.bind(on_keyboard=handle_keyboard)
    Clock.schedule_once(lambda dt: ajustar_layout(self, dt), 0.8)
    Clock.schedule_once(lambda dt: add_checkboxes_to_pdfs(self), 1)

def ajustar_layout(self, dt):
    try:
        if hasattr(self.file_manager, 'container'):
            self.file_manager.container.size_hint_x = 0.65 
    except Exception as e:
        print("Erro ao ajustar layout do FileManager:", e)

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

def select_pdf_file(self, paths):
    if isinstance(paths, str):
        paths = [paths]
    novos_pdfs = []
    for path in paths:
        if path.lower().endswith(".pdf") and path not in self.selected_files:
            self.selected_files.append(path)
            novos_pdfs.append(path)
            print(f"PDF adicionado: {path}")
        else:
            print("Arquivo já selecionado ou não é um PDF.")
    if novos_pdfs:
        finalize_selection(self)

def finalize_selection(self):
    from modules.leitorpdf import extrair_coordenadas_pdf, gerar_kml

    close_file_manager(self)

    if not self.selected_files:
        show_dialog(self, "Nenhum arquivo selecionado", "Selecione ao menos um PDF.")
        return

    for path in self.selected_files:
        try:
            coords, tipo, origens = extrair_coordenadas_pdf(path)
            if coords:
                gerar_kml(coords, path, tipo)
                show_dialog(self, "Sucesso", f"Coordenadas extraídas e KML gerado:\n{path}")
            else:
                show_dialog(self, "Erro", f"Não foi possível extrair coordenadas de:\n{path}")
        except Exception as e:
            show_dialog(self, "Erro", f"Erro ao processar o arquivo:\n{path}\n{str(e)}")

    self.selected_files.clear()

def show_dialog(self, title, text):
    dialog = MDDialog(
        MDDialogHeadlineText(text=title),
        MDDialogSupportingText(text=text),
        MDDialogButtonContainer(
            MDButton(
                MDButtonText(text="OK"),
                on_release=lambda x: dialog.dismiss()
            )
        )
    )
    dialog.open()