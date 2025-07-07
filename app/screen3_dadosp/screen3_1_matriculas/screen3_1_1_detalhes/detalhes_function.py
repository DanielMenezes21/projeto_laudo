def go_back(self):
    self.manager.current_screen.manager.current = "matricula"

def coletar_checkboxes(matricula_detalhe_screen):
    """
    Recebe uma instância de MatriculaDetalheScreen e retorna os estados das checkboxes.
    """
    checkboxes_af = {letra: cb.active for letra, cb in matricula_detalhe_screen.checkboxes_af.items()}
    checkboxes_pares = {par: cb.active for par, cb in matricula_detalhe_screen.checkboxes_pares.items()}
    checkboxes_superficie = {nome: cb.active for nome, cb in matricula_detalhe_screen.checkboxes_superficie.items()}
    return {
        "checkboxes_af": checkboxes_af,
        "checkboxes_pares": checkboxes_pares,
        "checkboxes_superficie": checkboxes_superficie,
        "poligono_regular": getattr(matricula_detalhe_screen, "checkbox_regular", None) and matricula_detalhe_screen.checkbox_regular.active,
        "poligono_irregular": getattr(matricula_detalhe_screen, "checkbox_irregular", None) and matricula_detalhe_screen.checkbox_irregular.active,
        "observacoes_imovel": getattr(matricula_detalhe_screen, "campo_observacoes", None).text if hasattr(matricula_detalhe_screen, "campo_observacoes") else ""
    }

def ir_para_parecer(self):
    nome_matricula = getattr(self, "nome_matricula", None)
    indice = getattr(self, "indice_matricula_atual", None)
    nome_tela_parecer = f"parecer_{nome_matricula}"

    if not self.manager.has_screen(nome_tela_parecer):
        from app.screen3_dadosp.screen3_1_matriculas.screen3_1_2_parecer.matricula_detalhe_parecer import MatriculaParecerScreen
        tela_parecer = MatriculaParecerScreen(
            nome_matricula,
            detalhes_screen=self,
            lista_dados_matriculas=self.lista_dados_matriculas,  
            indice_matricula_atual=indice,
            name=nome_tela_parecer
        )
        self.manager.add_widget(tela_parecer)
    self.manager.current = nome_tela_parecer