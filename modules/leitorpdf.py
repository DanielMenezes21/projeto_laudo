import re
import os
import fitz  
import io
import numpy as np
from PIL import Image
import geopandas as gpd
from shapely.geometry import Point, Polygon, MultiPolygon
import simplekml
import easyocr
from shapely.validation import explain_validity

ocr_reader = easyocr.Reader(['pt'])

def safe_convert(num_str):
    if isinstance(num_str, (float, int)):
        return float(num_str)
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
    
def detectar_zona_utm(easting):
    if 300000 <= easting < 400000:
        return "EPSG:31981"  # Zona 21S
    elif 400000 <= easting < 500000:
        return "EPSG:31982"  # Zona 22S
    elif 500000 <= easting < 600000:
        return "EPSG:31983"  # Zona 23S
    elif 600000 <= easting < 700000:
        return "EPSG:31984"  # Zona 24S
    return "EPSG:31982"

def pdf_tem_imagem(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    for pagina in doc:
        if pagina.get_images(full=True):
            return True
    return False

def extrair_coordenadas_pdf(caminho_pdf, forcar_ocr=False):
    import re
    doc = fitz.open(caminho_pdf)
    coordenadas = []
    coordenadas_originais = []
    tipo = None
    origens = []
    texto_total = ""

    for pagina in doc:
        pagina_texto = pagina.get_text().strip()
        texto_total += " " + pagina_texto

    pattern = r"""
    N[\-=]\s*((?=[\d\.,]*\d)[\d\.,]+)\s*.*?E[\-=']\s*((?=[\d\.,]*\d)[\d\.,]+)|
    Longitude[:=]?\s*([~\-+]?\d+)[°º*]?\s*(\d+)?[\'’′`´m]?\s*(\d+(?:[\.,]\d+)?)?["”″s]?\s*[,;]?\s*
    Latitude[:=]?\s*([~\-+]?\d+)[°º*]?\s*(\d+)?[\'’′`´m]?\s*(\d+(?:[\.,]\d+)?)?["”″s]?          
    """
    matches = re.findall(pattern, texto_total, re.VERBOSE)

    for match in matches:
        if match[0] and match[1]:
            tipo = "utm"
            n_raw, e_raw = match[0], match[1]
            n = n_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '').strip() if n_raw else "0"
            e = e_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '').strip() if e_raw else "0"
            try:
                n = safe_convert(n)
                e = safe_convert(e)
                coordenadas.append((float(f"{e:.3f}"), float(f"{n:.3f}")))
                coordenadas_originais.append((e_raw, n_raw))
                coordenadas_originais
                origens.append("utm-texto")
            except ValueError as e:
                print(f"Erro ao converter coordenadas UTM: {e}")
        elif match[2] and match[5]:
            tipo = "wgs84"
            lon_grau, lon_min, lon_sec, lat_grau, lat_min, lat_sec = match[2:8]
            lon_grau = lon_grau.strip() if lon_grau else "0"
            lon_min = lon_min.strip() if lon_min else "0"
            lon_sec = lon_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lon_sec else "0"
            lat_grau = lat_grau.strip() if lat_grau else "0"
            lat_min = lat_min.strip() if lat_min else "0"
            lat_sec = lat_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lat_sec else "0"
            if not lon_grau.startswith('-'):
                lon_grau = '-' + lon_grau
            if not lat_grau.startswith('-'):
                lat_grau = '-' + lat_grau
            longitude = dms_para_decimal(lon_grau, lon_min, lon_sec)
            latitude = dms_para_decimal(lat_grau, lat_min, lat_sec)
            if latitude is not None and longitude is not None and coordenada_valida(latitude, longitude):
                coordenadas.append((latitude, longitude))
                coordenadas_originais.append((lat_grau, lon_grau))
                origens.append("wgs84-texto")

    if not coordenadas or forcar_ocr:
        print("Nenhuma coordenada encontrada no texto digital. Tentando OCR nas imagens...")
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=1300)
            img_bytes = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            img_np = np.array(img)
            results = ocr_reader.readtext(img_np)
            ocr_texto = " ".join([txt for (_, txt, _) in results])
            matches = re.findall(pattern, ocr_texto, re.VERBOSE)
            arquivo_texto = f"{caminho_pdf}.txt"
            with open(arquivo_texto, "a", encoding="utf-8") as f:
                f.write(ocr_texto + "\n")
            for match in matches:
                if match[0] and match[1]:
                    tipo = "utm"
                    n_raw, e_raw = match[0], match[1]
                    n = n_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '').strip() if n_raw else "0"
                    e = e_raw.replace('.', '').replace(',', '.').replace('~', '').replace('=', '').strip() if e_raw else "0"
                    try:
                        if not n.startswith('-'):
                            n = '-' + n
                        if not e.startswith('-'):
                            e = '-' + e
                        n = safe_convert(n)
                        e = safe_convert(e)
                        coordenadas.append((float(f"{e:.3f}"), float(f"{n:.3f}")))
                        coordenadas_originais.append((e_raw, n_raw))
                        origens.append("utm-ocr")
                    except ValueError as e:
                        print(f"Erro ao converter coordenadas UTM: {e}")
                elif match[2] and match[5]:
                    tipo = "wgs84"
                    lon_grau, lon_min, lon_sec, lat_grau, lat_min, lat_sec = match[2:8]
                    lon_grau = lon_grau.strip() if lon_grau else "0"
                    lon_min = lon_min.strip() if lon_min else "0"
                    lon_sec = lon_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lon_sec else "0"
                    lat_grau = lat_grau.strip() if lat_grau else "0"
                    lat_min = lat_min.strip() if lat_min else "0"
                    lat_sec = lat_sec.strip().replace(',', '.').replace('~', '').replace('=', '') if lat_sec else "0"
                    if not lon_grau.startswith('-'):
                        lon_grau = '-' + lon_grau
                    if not lat_grau.startswith('-'):
                        lat_grau = '-' + lat_grau
                    longitude = dms_para_decimal(lon_grau, lon_min, lon_sec)
                    latitude = dms_para_decimal(lat_grau, lat_min, lat_sec)
                    if latitude is not None and longitude is not None and coordenada_valida(latitude, longitude):
                        coordenadas.append((latitude, longitude))
                        coordenadas_originais.append((lat_grau, lat_min, lat_sec, lon_grau, lon_min, lon_sec))
                        origens.append("wgs84-ocr")

    tipo = "utm" if any(o.startswith("utm") for o in origens) else "wgs84" if coordenadas else None

    arquivo_texto = f"{caminho_pdf}_coordenadas.txt"
    with open(arquivo_texto, "w", encoding="utf-8") as f:
        for idx, ((lat, lon), orig) in enumerate(zip(coordenadas, coordenadas_originais), 1):
            original_str = ", ".join(str(x) for x in orig)
            f.write(f"{idx}. {lat:.8f}, {lon:.8f} | original: {original_str}\n")
    return coordenadas, tipo, origens

def remove_pontos_duplicados(pontos):
    pontos_sem_duplicados = []
    for pt in pontos:
        if not pontos_sem_duplicados or pt != pontos_sem_duplicados[-1]:
            pontos_sem_duplicados.append(pt)
    return pontos_sem_duplicados

def gerar_kml(coordenadas, caminho_pdf, tipo):
    if tipo == "utm":
        pontos_utm = [(float(e), float(n)) for e, n in coordenadas]
        zona_utm = detectar_zona_utm(pontos_utm[0][0])
        gdf = gpd.GeoDataFrame(geometry=[Point(e, n) for e, n in pontos_utm], crs=zona_utm)
        gdf_wgs84 = gdf.to_crs(epsg=4326)
        pontos = [(pt.x, pt.y) for pt in gdf_wgs84.geometry]
    else:
        pontos = [(float(lon), float(lat)) for lat, lon in coordenadas]

    if pontos and (pontos[0] != pontos[-1]):
        pontos.append(pontos[0])

    poligono = Polygon(pontos)
    if not poligono.is_valid:
        motivo = explain_validity(poligono)
        print(f"Polígono inválido detectado. Motivo: {motivo}")
        if "at" in motivo:
            try:
                ponto_str = motivo.split("at")[1].strip()
                print(f"Ponto problemático: {ponto_str}")
            except Exception:
                pass
        print("Tentando corrigir...")
        poligono = poligono.buffer(0)
        if not poligono.is_valid:
            novo_motivo = explain_validity(poligono)
            print(f"Falha ao corrigir o polígono. Ainda inválido. Motivo: {novo_motivo}")
        else:
            print("Polígono corrigido com sucesso!")

    if isinstance(poligono, MultiPolygon):
        print("MultiPolygon detectado. Usando apenas o maior polígono.")
        poligono = max(poligono.geoms, key=lambda p: p.area)  

    pontos_corrigidos = list(poligono.exterior.coords)

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
    caminho_pdf3 = r"H:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL\\05. MAIO\\Processo nº 123456789 - JOELSON SOUSA JUNIOR\\enhance_pdf.PDF"

    coords, tipo1, origens1 = extrair_coordenadas_pdf(caminho_pdf)
    coords2, tipo2, origens3= extrair_coordenadas_pdf(caminho_pdf2)
    coords3, tipo3, origens2= extrair_coordenadas_pdf(caminho_pdf3)
    
    if coords:
        gerar_kml(coords, caminho_pdf, tipo1)
    if coords2:
        gerar_kml(coords2, caminho_pdf2, tipo2)
    if coords3:
        gerar_kml(coords3, caminho_pdf3, tipo3)
