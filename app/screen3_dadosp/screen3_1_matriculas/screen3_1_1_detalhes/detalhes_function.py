from kivymd.uix.filemanager import MDFileManager
import os

def go_back(self):
    self.manager.current_screen.manager.current = "matricula"

def coletar_dados(matricula_detalhe_screen):
    """
    Recebe uma instância de MatriculaDetalheScreen e retorna os estados das checkboxes.
    """
    checkboxes_af = {letra: cb.active for letra, cb in matricula_detalhe_screen.checkboxes_af.items()}
    checkboxes_pares = {par: cb.active for par, cb in matricula_detalhe_screen.checkboxes_pares.items()}
    checkboxes_superficie = {nome: cb.active for nome, cb in matricula_detalhe_screen.checkboxes_superficie.items()}
    print(type(coletar_dados))
    return {
        "checkboxes_af": checkboxes_af,
        "checkboxes_pares": checkboxes_pares,
        "checkboxes_superficie": checkboxes_superficie,
        "poligono_regular": getattr(matricula_detalhe_screen, "checkbox_regular", None) and matricula_detalhe_screen.checkbox_regular.active,
        "poligono_irregular": getattr(matricula_detalhe_screen, "checkbox_irregular", None) and matricula_detalhe_screen.checkbox_irregular.active,
        "observacoes_imovel": getattr(matricula_detalhe_screen, "campo_observacoes", None).text if hasattr(matricula_detalhe_screen, "campo_observacoes") else "",
        "p_reserva": getattr(matricula_detalhe_screen, "p_reserva", None).text if hasattr(matricula_detalhe_screen, "p_reserva") else "",
        "area_reserva": getattr(matricula_detalhe_screen, "area_reserva", None).text if hasattr(matricula_detalhe_screen, "area_reserva") else "",
        "p_app": getattr(matricula_detalhe_screen, "p_app", None).text if hasattr(matricula_detalhe_screen, "p_app") else "",
        "a_app": getattr(matricula_detalhe_screen, "a_app", None).text if hasattr(matricula_detalhe_screen, "a_app") else "",
        "atividade_imovel": getattr(matricula_detalhe_screen, "atividade_imovel", None).text if hasattr(matricula_detalhe_screen, "atividade_imovel") else "",
        "atividade_potencial": getattr(matricula_detalhe_screen, "atividade_potencial", None).text if hasattr(matricula_detalhe_screen, "atividade_potencial") else "",
        "description_img_one": getattr(matricula_detalhe_screen, "description_img_one", None).text if hasattr(matricula_detalhe_screen, "description_img_one") else "",
        "description_img_two": getattr(matricula_detalhe_screen, "description_img_two", None).text if hasattr(matricula_detalhe_screen, "description_img_two") else "",
        "img_one": getattr(matricula_detalhe_screen, "img_one", "") if hasattr(matricula_detalhe_screen, "img_one") else "",
        "img_two": getattr(matricula_detalhe_screen, "img_two", "") if hasattr(matricula_detalhe_screen, "img_two") else ""
    }


def _atualizar_dados_apos_salvar(self, instance, indice):
    if 0 <= indice < len(self.lista_dados_matriculas):
        print(f"Dados atualizados para matrícula {indice}")

def ir_para_parecer(self):
    nome_matricula = getattr(self, "nome_matricula", "")
    indice = getattr(self, "indice_matricula", None)
    
    if indice is None:
        print("❌ Dados insuficientes para abrir parecer")
        return

    nome_tela = f"parecer_{nome_matricula}_{indice}"  

    if not self.manager.has_screen(nome_tela):
        from app.screen3_dadosp.screen3_1_matriculas.screen3_1_2_parecer.matricula_detalhe_parecer import MatriculaParecerScreen
        
        if indice >= len(self.lista_dados_matriculas):
            self.lista_dados_matriculas.append({})
            
        tela = MatriculaParecerScreen(
            nome_matricula=nome_matricula,
            detalhes_screen=self,
            lista_dados_matriculas=self.lista_dados_matriculas,
            indice_matricula=indice,
            name=nome_tela
        )
        self.manager.add_widget(tela)
    
    self.manager.current = nome_tela

def selecionar_imagem(self, instance):
    """
    Abre o gerenciador de arquivos para selecionar uma imagem, distinguindo entre img_one e img_two.
    """
    self.file_manager = MDFileManager(
        exit_manager=lambda *args: exit_manager(self, *args),
        select_path=lambda path: select_path(self, path),
        preview=True
    )
    if instance == self.image_one_selection:
        self.current_image_selection = "img_one"
    elif instance == self.image_two_selection:
        self.current_image_selection = "img_two"
    self.file_manager.show(self.current_path if hasattr(self, "current_path") else os.path.expanduser("~"))

def exit_manager(self, *args):
    """
    Fecha o gerenciador de arquivos.
    """
    if hasattr(self, "file_manager") and self.file_manager:
        self.file_manager.close()

def select_path(self, path):
    """
    Recebe o caminho do arquivo selecionado e atualiza img_one ou img_two.
    """
    if hasattr(self, "current_image_selection"):
        if self.current_image_selection == "img_one":
            self.img_one = path
            print(f"🖼️ Primeira imagem selecionada: {path}")
        elif self.current_image_selection == "img_two":
            self.img_two = path
            print(f"🖼️ Segunda imagem selecionada: {path}")
    exit_manager(self)
