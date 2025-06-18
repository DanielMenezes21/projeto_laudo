from kivymd.uix.list import MDList, MDListItem, MDListItemHeadlineText, MDListItemTrailingCheckbox
import os
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.metrics import dp
import geopandas as gpd

def go_back(self, *args):
    self.manager.current_screen.manager.current = "main"

def go_next(self, *args):
    self.manager.current_screen.manager.current = "manager"

def load_directory(self, *args):
    self.file_list.clear_widgets()

    if self.current_path != self.root_path:
        voltar_item = MDListItem(
            on_release=lambda x: go_up(self) 
        )
        voltar_item.add_widget(MDListItemHeadlineText(text=".. (voltar)"))
        self.file_list.add_widget(voltar_item)

    try:
        itens = os.listdir(self.current_path)
    except FileNotFoundError:
        itens = []

    for nome in sorted(itens):
        caminho = os.path.join(self.current_path, nome)

        if os.path.isdir(caminho):
            item = MDListItem(
                on_release=lambda x, p=caminho: entrar_em_pasta(self, p)
            )
            item.add_widget(MDListItemHeadlineText(text=f"[DIR] {nome}"))
            self.file_list.add_widget(item)

        elif nome.lower().endswith(".kml"):
            item = MDListItem(
                on_release=lambda x, f=caminho: on_file_selected(self, f)
            )
            item.add_widget(MDListItemHeadlineText(text=nome))
            self.file_list.add_widget(item)

def entrar_em_pasta(self, pasta):
    self.current_path = pasta
    load_directory(self)

def go_up(self, *args):
    self.current_path = os.path.dirname(self.current_path)
    load_directory(self)

def on_file_selected(self, caminho):
    try:
        gdf = gpd.read_file(caminho, driver="KML")
        shp_path = os.path.splitext(caminho)[0] + ".shp"

        gdf.to_file(shp_path, driver="ESRI Shapefile")

        MDSnackbar(
            MDSnackbarText(text=f"Arquivo convertido para SHP:\n{shp_path}"),
            y=dp(24),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.8,
        ).open()
    except Exception as e:
        MDSnackbar(
            MDSnackbarText(text=f"Erro na conversão:\n{str(e)}"),
            y=dp(24),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.8,
        ).open()