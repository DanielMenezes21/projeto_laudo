from kivymd.uix.pickers import MDModalDatePicker, MDModalInputDatePicker
from kivy.clock import Clock
from datetime import datetime

"""def show_modal_input_date_picker(self, *args):
    def on_ok(instance, value=None, date_range=None):
        print(f"on_ok called: value={value}, date_range={date_range}")
        if value:
            formatted_date = value.strftime('%d/%m/%Y')
            self.data_vistoria_text.text = formatted_date
            print(f"Data de vistoria selecionada: {formatted_date}")
        else:
            print("Nenhuma data foi selecionada ou digitada.")
        instance.dismiss()

    def on_edit(instance, *a):
        instance.dismiss()
        show_modal_date_picker(self)

    date_dialog = MDModalInputDatePicker(date_format="dd/mm/yyyy")
    date_dialog.bind(on_ok=on_ok, on_edit=on_edit)
    date_dialog.open()"""

def show_modal_date_picker(self, *args):
    """Abre o diálogo para selecionar a data."""
    def on_ok(instance_date_picker, *a):
        try:
            selected_date = instance_date_picker.get_date()[0]
            print(selected_date)
            formatted_date = selected_date.strftime('%d/%m/%Y')
            self.data_vistoria_text.text = formatted_date
            print(f"Data de vistoria selecionada: {formatted_date}")
        except Exception as e:
            print(f"Erro ao capturar data: {e}")
            print("Nenhuma data foi selecionada ou digitada.")
        instance_date_picker.dismiss()

    def on_edit(*args):
        date_dialog.dismiss()

    date_dialog = MDModalDatePicker()
    date_dialog.bind(on_ok=on_ok, on_edit=on_edit)
    date_dialog.open()

