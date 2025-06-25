import fitz  # PyMuPDF

def print_text_with_unicode(pdf_path):
    doc = fitz.open(pdf_path)
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        raw_text = page.get_text("text")
        
        print(f"\n=== Página {page_num + 1} ===")
        # Converte caracteres especiais para representação Unicode escape
        encoded_text = raw_text.encode('unicode-escape').decode('ascii')
        print(encoded_text)

# Exemplo de uso:
print_text_with_unicode("test\\2632-2620-2612-2172_TO-1720655-9F36D450A8BA4273BC1D685694E4F4DA.pdf")