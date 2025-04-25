import os
import fitz
import uuid

def extrair_paginas_como_imagens(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        if not doc:
            print(f"Erro ao abrir o PDF: {pdf_path}")
            return []

        imagens_salvas = []
        print(f"Total de páginas no PDF: {len(doc)}")

        for i, pagina in enumerate(doc):
            try:
                imagem = pagina.get_pixmap()
                nome_arquivo = f"pagina_{uuid.uuid4().hex}.png"
                caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)

                imagem.save(caminho_arquivo)
                imagens_salvas.append(caminho_arquivo)
                #print(f"Imagem da página {i+1} salva em: {caminho_arquivo}")
            
            except Exception as e:
                print(f"Erro ao extrair imagem da página {i+1}: {e}")

        return imagens_salvas

    except Exception as e:
        print(f"Erro ao abrir o PDF {pdf_path}: {e}")
        return []



