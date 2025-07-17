from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.label import MDLabel
import os

def add_checkboxes_to_pdfs(screen):
    if not hasattr(screen, "file_manager") or not screen.file_manager:
        return

    layout = screen.file_manager.children[0]

    for child in layout.children[:]:
        if hasattr(child, "text") and child.text.endswith(".pdf"):
            container = BoxLayout(orientation="horizontal", size_hint_y=None, height=40)
            
            label = MDLabel(text=os.path.basename(child.text), halign="left", valign="middle")
            checkbox = MDCheckbox(size_hint=(None, None), size=(40, 40))

            checkbox.path = child.text

            def on_checkbox_active(instance, value):
                if value:
                    if instance.path not in screen.selected_files:
                        screen.selected_files.append(instance.path)
                else:
                    if instance.path in screen.selected_files:
                        screen.selected_files.remove(instance.path)

            checkbox.bind(active=on_checkbox_active)

            container.add_widget(label)
            container.add_widget(checkbox)
            layout.add_widget(container)
