import os
import fitz
import uuid

def extrair_paginas_como_imagens(pdf_path):
    doc = fitz.open(pdf_path)
    imagens_salvas = []

    for i, pagina in enumerate(doc):
        imagem = pagina.get_pixmap()
        nome_arquivo = f"pagina_{uuid.uuid4().hex}.png"
        caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)
        imagem.save(caminho_arquivo)
        imagens_salvas.append(caminho_arquivo)

    return imagens_salvas
