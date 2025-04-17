import re
import os
import fitz  
import geopandas as gpd
from shapely.geometry import Polygon

def extrair_coordenadas_pdf(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto = ""
    for pagina in doc:
        texto += pagina.get_text()

    texto = texto.replace("\n", " ").replace("\xa0", " ")
    pattern = r"N\s*=\s*([\d\.]+,[\d]+).*?E\s*=\s*([\d\.]+,[\d]+)"
    matches = re.findall(pattern, texto)

    coordenadas = []
    for n_raw, e_raw in matches:
        n = float(n_raw.replace('.', '').replace(',', '.'))
        e = float(e_raw.replace('.', '').replace(',', '.'))
        coordenadas.append((n, e))

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
