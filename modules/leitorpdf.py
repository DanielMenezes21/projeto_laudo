import re
import os
import fitz  
import geopandas as gpd
from shapely.geometry import Polygon

def safe_convert(num_str):
    """
    Remove caracteres desnecessários e tenta converter a string para float.
    Assume que se há vírgula e ponto, o ponto é separador de milhares e a vírgula o decimal.
    """

    cleaned = num_str.strip()
    cleaned = re.sub(r'\s+', '', cleaned)
    if not re.search(r'\d', cleaned):
        raise ValueError("Nenhum dígito encontrado em '{}'".format(num_str))
    cleaned = re.sub(r'^[^\d]+', '', cleaned)
    cleaned = re.sub(r'[^\d]+$', '', cleaned)
    if ',' in cleaned and '.' in cleaned:
        cleaned = cleaned.replace('.', '').replace(',', '.')
    elif cleaned.count(',') > 1:
        cleaned = cleaned.replace(',', '')
    elif cleaned.count('.') > 1:
        partes = cleaned.split('.')
        cleaned = ''.join(partes[:-1]) + '.' + partes[-1]
    elif ',' in cleaned:
        cleaned = cleaned.replace(',', '.')
    return float(cleaned)

def extrair_coordenadas_pdf(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto = ""
    for pagina in doc:
        texto += pagina.get_text()

    texto = texto.replace("\n", " ").replace("\xa0", " ")
    pattern = r"N\s*=\s*((?=[\d\.,]*\d)[\d\.,]+).*?E\s*=\s*((?=[\d\.,]*\d)[\d\.,]+)"
    matches = re.findall(pattern, texto)

    coordenadas = []
    for n_raw, e_raw in matches:
        if n_raw.strip() and e_raw.strip():
            try:
                n = safe_convert(n_raw)
                e = safe_convert(e_raw)
                coordenadas.append((n, e))
                print(f"Coordenadas extraídas: {n}, {e}")
            except ValueError:
                print(f"Erro ao converter coordenadas: {n_raw}, {e_raw}")
    return coordenadas

def gerar_kml(coordenadas, caminho_pdf):
    base, _ = os.path.splitext(caminho_pdf)
    nome_arquivo = f"{base}.kml"

    pontos = [(e, n) for n, e in coordenadas]
    if pontos[0] != pontos[-1]:
        pontos.append(pontos[0])
    poligono = Polygon(pontos)

    gdf = gpd.GeoDataFrame(
        {'name': ['Fazenda Bandeirante']},
        geometry=[poligono],
        crs="EPSG:32722" 
    )

    gdf_wgs84 = gdf.to_crs(epsg=4326)
    gdf_wgs84.to_file(nome_arquivo, driver="KML")

    print(f"KML '{nome_arquivo}' criado com sucesso!")

if __name__ == "__main__":
    caminho_pdf = r"anexos\Processo nº 123456789 - JOELSON SOUSA JUNIOR\CERT_INTEIRO_TEOR_M.11173.pdf"
    coords = extrair_coordenadas_pdf(caminho_pdf)

    if coords:
        gerar_kml(coords, caminho_pdf)
    else:
        print("Nenhuma coordenada encontrada.")
