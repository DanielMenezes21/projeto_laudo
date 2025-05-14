import re
import os
import fitz  
import io
import numpy as np
from PIL import Image
import geopandas as gpd
from shapely.geometry import Point, Polygon
import simplekml
import easyocr

ocr_reader = easyocr.Reader(['pt'])

def safe_convert(num_str):
    cleaned = num_str.strip()
    cleaned = re.sub(r'\s+', '', cleaned)
    if not re.search(r'\d', cleaned):
        raise ValueError(f"Nenhum dígito encontrado em '{num_str}'")
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

def dms_para_decimal(grau, minuto, segundo):
    try:
        grau = float(str(grau).replace(',', '.')) if grau else 0
        minuto = float(str(minuto).replace(',', '.')) if minuto else 0
        segundo = float(str(segundo).replace(',', '.')) if segundo else 0

        valor = abs(grau) + (minuto / 60) + (segundo / 3600)

        return -valor if grau < 0 else valor
    except Exception as e:
        print(f"Erro ao converter DMS para decimal: {e}")
        return None

def coordenada_valida(lat, lon):
    try:
        lat = float(lat)
        lon = float(lon)
        return -90 <= lat <= 90 and -180 <= lon <= 180
    except Exception:
        return False

def extrair_coordenadas_pdf(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto = ""

    for pagina in doc:
        texto_pagina = ""

        pagina_texto = pagina.get_text().strip()
        if pagina_texto:
            texto_pagina += pagina_texto

        pix = pagina.get_pixmap(dpi=300) 
        img_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img_np = np.array(img)

        results = ocr_reader.readtext(img_np)
        ocr_texto = " ".join([txt for (_, txt, _) in results])
        if ocr_texto:
            ocr_texto = ocr_texto.replace("\n", " ").replace("\xa0", " ")
            ocr_texto = re.sub(r'\s+', ' ', ocr_texto).strip()
            if len(ocr_texto) > 10:
                #print(f"Texto OCR extraído: {ocr_texto}")
                base, _ = os.path.splitext(caminho_pdf)
                nome_arquivo = f"{base}_ocr.txt"
                with open(nome_arquivo, "a", encoding="utf-8") as f:
                    f.write(ocr_texto + "\n")
                
            else:
                print("Texto OCR extraído é muito curto ou inválido.")

        texto += texto_pagina + " " + ocr_texto + " "

    texto = texto.replace("\n", " ").replace("\xa0", " ")
    pattern = r"""
    N\s*=\s*((?=[\d\.,]*\d)[\d\.,]+)\s*.*?E\s*=\s*((?=[\d\.,]*\d)[\d\.,]+)|  
    Longitude:\s*([-+]?\d+)[°*\s]?\s*(\d+)?[\'\s]?\s*(\d+(?:[\.,]\d+)?)?["\s]?\s*,?\s*  
    Latitude:\s*([-+]?\d+)[°*\s]?\s*(\d+)?[\'\s]?\s*(\d+(?:[\.,]\d+)?)?["\s]?          
    """
    matches = re.findall(pattern, texto, re.VERBOSE)

    coordenadas = []
    tipo = None 

    for match in matches:
        if match[0] and match[1]:
            tipo = "utm"
            n_raw, e_raw = match[0], match[1]
            n = n_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '') if ',' in n_raw else n_raw
            e = e_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '') if ',' in e_raw else e_raw
            coordenadas.append((n, e))
        elif match[2] and match[5]:
            tipo = "wgs84"
            lon_grau, lon_min, lon_sec, lat_grau, lat_min, lat_sec = match[2:8]

            lon_grau = lon_grau.strip() if lon_grau else "0"
            lon_min = lon_min.strip() if lon_min else "0"
            lon_sec = lon_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lon_sec else "0"

            lat_grau = lat_grau.strip() if lat_grau else "0"
            lat_min = lat_min.strip() if lat_min else "0"
            lat_sec = lat_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lat_sec else "0"

            print(f"Longitude capturada: {lon_grau}° {lon_min}' {lon_sec}\"")
            print(f"Latitude capturada: {lat_grau}° {lat_min}' {lat_sec}\"")

            longitude = dms_para_decimal(lon_grau, lon_min, lon_sec)
            latitude = dms_para_decimal(lat_grau, lat_min, lat_sec)
            for i in latitude, longitude:
                if latitude is not None and longitude is not None:
                    arquivo_coordenadas = f"{caminho_pdf}_coordenadas.txt"
                    with open(arquivo_coordenadas, "a", encoding="utf-8") as f:
                        f.write(f"Latitude: {latitude}, Longitude: {longitude}\n")


            if latitude is not None and longitude is not None:
                if coordenada_valida(latitude, longitude):
                    coordenadas.append((latitude, longitude))
                    #print(f"Coordenadas válidas: Latitude: {latitude}, Longitude: {longitude}")
                    arquivo_coordenadas = f"{caminho_pdf}_coordenadas.txt"
                    with open(arquivo_coordenadas, "a", encoding="utf-8") as f:
                        f.write(f"Latitude: {latitude}, Longitude: {longitude}\n")
                else:
                    print(f"Coordenada INVÁLIDA descartada: Latitude: {latitude}, Longitude: {longitude}")

    if not tipo:
        tipo = "utm" if coordenadas else None

    return coordenadas, tipo

def gerar_kml(coordenadas, caminho_pdf, tipo):
    if tipo == "utm":
        pontos_utm = [(float(e), float(n)) for n, e in coordenadas]
        gdf = gpd.GeoDataFrame(geometry=[Point(p) for p in pontos_utm], crs="EPSG:31982")
        gdf_wgs84 = gdf.to_crs(epsg=4326)
        pontos = [(pt.x, pt.y) for pt in gdf_wgs84.geometry]
    else:
        pontos = [(float(lon), float(lat)) for lat, lon in coordenadas]

    # Fechar o polígono se necessário
    if pontos[0] != pontos[-1]:
        pontos.append(pontos[0])

    # Criar o polígono e verificar sua validade
    poligono = Polygon(pontos)
    if not poligono.is_valid:
        print("Polígono inválido detectado. Tentando corrigir...")
        poligono = poligono.buffer(0)  # Corrige o polígono automaticamente

    # Converter o polígono corrigido de volta para pontos
    pontos_corrigidos = list(poligono.exterior.coords)

    # Criar o KML
    kml = simplekml.Kml()
    poligono_kml = kml.newpolygon(name="Área delimitada", outerboundaryis=pontos_corrigidos)
    poligono_kml.style.linestyle.width = 2
    poligono_kml.style.linestyle.color = simplekml.Color.white
    poligono_kml.style.polystyle.color = simplekml.Color.changealphaint(100, simplekml.Color.green)

    base, _ = os.path.splitext(caminho_pdf)
    nome_arquivo = f"{base}.kml"
    kml.save(nome_arquivo)
    print(f"KML '{nome_arquivo}' criado com sucesso!")

if __name__ == "__main__":
    caminho_pdf = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL\\05. MAIO\\Processo nº 123456789 - JOELSON SOUSA JUNIOR\\CERT_INTEIRO_TEOR_M.11173.pdf"
    caminho_pdf2 = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL\\05. MAIO\\Processo nº 123456789 - JOELSON SOUSA JUNIOR\\CIT.PDF"
    caminho_pdf3 = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL\\05. MAIO\\Processo nº 123456789 - JOELSON SOUSA JUNIOR\\CIT2.PDF"

    coords, tipo1 = extrair_coordenadas_pdf(caminho_pdf)
    coords2, tipo2= extrair_coordenadas_pdf(caminho_pdf2)
    coords3, tipo3= extrair_coordenadas_pdf(caminho_pdf3)
    
    if coords and coords2 and coords3:
        gerar_kml(coords, caminho_pdf, tipo1)
        gerar_kml(coords2, caminho_pdf2, tipo2)
        gerar_kml(coords3, caminho_pdf3, tipo3)
