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
        "observacoes_imovel": getattr(matricula_detalhe_screen, "campo_observacoes", None).text if hasattr(matricula_detalhe_screen, "campo_observacoes") else "",
        "p_reserva": getattr(matricula_detalhe_screen, "p_reserva", None).text if hasattr(matricula_detalhe_screen, "p_reserva") else "",
        "area_reserva": getattr(matricula_detalhe_screen, "area_reserva", None).text if hasattr(matricula_detalhe_screen, "area_reserva") else "",
        "p_app": getattr(matricula_detalhe_screen, "p_app", None).text if hasattr(matricula_detalhe_screen, "p_app") else "",
        "a_app": getattr(matricula_detalhe_screen, "a_app", None).text if hasattr(matricula_detalhe_screen, "a_app") else ""
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