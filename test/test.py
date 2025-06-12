import re
import pandas as pd
import geopandas as gpd
from shapely.geometry import Polygon
from PyPDF2 import PdfReader
from simplekml import Kml  

def extract_coordinates_from_pdf(pdf_path):
    """
    Extrai coordenadas de um arquivo PDF com o formato específico fornecido.
    Retorna uma lista de tuplas (longitude, latitude) em graus decimais.
    """
    coordinates = []
    reader = PdfReader(pdf_path)
    
    for page in reader.pages:
        text = page.extract_text()
        matches = re.finditer(
            r"Latitude:\s*(-?\d{2})°(\d{2})'(\d{2})''(\d+)''\s*Longitude:\s*(-?\d{2})°(\d{2})'(\d{2})''(\d+)''",
            text
        )
        for match in matches:
            lat_deg, lat_min, lat_sec, lat_mil, lon_deg, lon_min, lon_sec, lon_mil = match.groups()
            lat = dms_to_dd(lat_deg, lat_min, lat_sec, lat_mil)
            lon = dms_to_dd(lon_deg, lon_min, lon_sec, lon_mil)
            coordinates.append((lon, lat))  
    
    return coordinates

def dms_to_dd(degrees, minutes, seconds, milliseconds):
    dd = float(degrees) + float(minutes)/60 + (float(seconds) + float(milliseconds)/1000)/3600
    return dd

def create_shapefile(coordinates, output_path):
    if not coordinates:
        raise ValueError("A lista de coordenadas está vazia. Verifique se o PDF foi lido corretamente.")
    if coordinates[0] != coordinates[-1]:
        coordinates.append(coordinates[0])
    polygon = Polygon(coordinates)
    gdf = gpd.GeoDataFrame(geometry=[polygon], crs="EPSG:4326")
    gdf.to_file(output_path, encoding='utf-8')

def save_coordinates_txt(coordinates, output_txt):
    with open(output_txt, "w", encoding="utf-8") as f:
        for lon, lat in coordinates:
            f.write(f"{lon},{lat}\n")

def create_kml(coordinates, output_kml):
    kml = Kml()
    pol = kml.newpolygon(name="Polígono", outerboundaryis=[(lon, lat) for lon, lat in coordinates])
    pol.style.linestyle.width = 2
    pol.style.linestyle.color = 'ff0000ff' 
    pol.style.polystyle.color = '7d00ff00'  
    kml.save(output_kml)

def remove_duplicates(coords):
    unique = []
    for c in coords:
        if c not in unique:
            unique.append(c)
    if unique and unique[0] != unique[-1]:
        unique.append(unique[0])
    return unique

if __name__ == "__main__":
    pdf_path = "C:\\Users\\DESKTOP\\Downloads\\FAZER_SHAPE.pdf"
    output_shp = "poligono_shape.shp"
    output_txt = "poligono_shape.txt"
    output_kml = "poligono_shape.kml"

    coords = extract_coordinates_from_pdf(pdf_path)

    coords = remove_duplicates(coords)

    print("Coordenadas extraídas:")
    for lon, lat in coords:
        print(f"Longitude: {lon}, Latitude: {lat}")

    create_shapefile(coords, output_shp)
    save_coordinates_txt(coords, output_txt)
    create_kml(coords, output_kml)
    
    print(f"Shapefile criado com sucesso: {output_shp}")
    print(f"TXT criado com sucesso: {output_txt}")
    print(f"KML criado com sucesso: {output_kml}")