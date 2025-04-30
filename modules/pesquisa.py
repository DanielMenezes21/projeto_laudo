import wikipedia
wikipedia.set_lang("pt")  

def buscar_descricao_cidade(nome_cidade):
    try:
        resumo = wikipedia.summary(nome_cidade, sentences=8)  # Retorna as 3 primeiras frases
        return resumo
    except wikipedia.exceptions.DisambiguationError as e:
        return f"⚠️ Muitos resultados encontrados. Seja mais específico: {e.options[:5]}"
    except wikipedia.exceptions.PageError:
        return f"❌ Página não encontrada para '{nome_cidade}'."
    except Exception as e:
        return f"❌ Erro inesperado: {str(e)}"