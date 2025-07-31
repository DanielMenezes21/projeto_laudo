import win32com.client
import pythoncom

def is_word_available():
    try:
        # Tenta obter uma instância ativa do Word (se estiver aberto)
        word = win32com.client.GetActiveObject("Word.Application")
    except pythoncom.com_error:
        # Se não estiver aberto, retorna False
        return False

    try:
        # Verifica se a aplicação está visível e não está em modo "não respondendo"
        if word.Visible:
            # Também verifica se o status não é de bloqueio
            # Infelizmente o Word não expõe um método direto para "disponibilidade", mas podemos tentar usar:
            # Por exemplo, se o Word estiver bloqueado, comandos disparam exceções.

            # Tentativa simples: checar se podemos acessar um atributo básico
            _ = word.Ready  # Algumas versões do Word têm essa propriedade

            # Se chegou até aqui, está disponível
            return True
        else:
            # Se não está visível, pode ser uma instância oculta, mas vamos assumir indisponível
            return False
    except Exception:
        # Se deu erro acessando propriedades, significa que o Word não está disponível
        return False


if __name__ == "__main__":
    if is_word_available():
        print("O Word está aberto e disponível para edição.")
    else:
        print("O Word não está aberto ou não está disponível.")
