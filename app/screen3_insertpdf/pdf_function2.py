import os
import fitz
import uuid

def extrair_paginas_como_imagens(pdf_path):
    try:
        # Abrindo o PDF
        doc = fitz.open(pdf_path)
        if not doc:
            print(f"Erro ao abrir o PDF: {pdf_path}")
            return []

        imagens_salvas = []
        print(f"Total de páginas no PDF: {len(doc)}")

        # Iterando sobre as páginas do PDF
        for i, pagina in enumerate(doc):
            try:
                # Gerando a imagem da página
                imagem = pagina.get_pixmap()
                nome_arquivo = f"pagina_{uuid.uuid4().hex}.png"
                caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)

                # Salvando a imagem
                imagem.save(caminho_arquivo)
                imagens_salvas.append(caminho_arquivo)
                print(f"Imagem da página {i+1} salva em: {caminho_arquivo}")
            
            except Exception as e:
                print(f"Erro ao extrair imagem da página {i+1}: {e}")

        return imagens_salvas

    except Exception as e:
        print(f"Erro ao abrir o PDF {pdf_path}: {e}")
        return []



