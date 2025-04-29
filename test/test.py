import geopandas as gpd
import math
import requests
from PIL import Image
from io import BytesIO
import os

# 1. CONFIGURAÇÕES -------------------------------

# CAMINHO DO SEU ARQUIVO .shp
CAMINHO_ARQUIVO = r"anexos\Processo nº 123456789 - JOELSON SOUSA JUNIOR\CERT_INTEIRO_TEOR_M.11173.shp"  # <<< AJUSTAR AQUI

# DATA DA IMAGEM QUE QUER PEGAR
DATA_IMAGEM = "2024-04-26"

# CAMADA GIBS QUE VOCÊ QUER
CAMADA = "VIIRS_SNPP_CorrectedReflectance_TrueColor"

# Nível de zoom
ZOOM_LEVEL = 9  # <<< Recomendo aumentar (tipo 10, 11 ou 12) para áreas pequenas

# Pasta para salvar imagens baixadas
PASTA_IMAGENS = "imagens_tiles"

# 2. FUNÇÕES AUXILIARES ---------------------------

def lonlat_to_tilexy_gibs(lon, lat, zoom):
    """Converte longitude/latitude para coordenadas X/Y no GIBS EPSG:4326"""
    num_tiles = 2 ** zoom
    tile_x = int((lon + 180) / 360 * num_tiles)
    tile_y = int((90 - lat) / 180 * num_tiles)
    return tile_x, tile_y

def montar_url_tile(layer, date, tile_matrix_set, zoom, tile_row, tile_col):
    """Monta a URL para pegar um tile do GIBS"""
    return f"https://gibs.earthdata.nasa.gov/wmts/epsg4326/best/{layer}/default/{date}/{tile_matrix_set}/{zoom}/{tile_row}/{tile_col}.png"

# 3. LER O SHAPEFILE E PEGAR BOUNDING BOX ----------------

print(f"Lendo arquivo: {CAMINHO_ARQUIVO}")
gdf = gpd.read_file(CAMINHO_ARQUIVO)

# Bounding box (minx, miny, maxx, maxy)
minx, miny, maxx, maxy = gdf.total_bounds
print(f"Bounding box: {minx}, {miny}, {maxx}, {maxy}")

# 4. CONVERTER BBOX PARA TILES -----------------------

x_min, y_min = lonlat_to_tilexy_gibs(minx, miny, ZOOM_LEVEL)
x_max, y_max = lonlat_to_tilexy_gibs(maxx, maxy, ZOOM_LEVEL)

print(f"Tile range: x {x_min} -> {x_max}, y {y_min} -> {y_max}")

# Corrige ordem
x_start, x_end = sorted([x_min, x_max])
y_start, y_end = sorted([y_min, y_max])

# 5. BAIXAR TODOS OS TILES ---------------------------

os.makedirs(PASTA_IMAGENS, exist_ok=True)

tiles_baixados = []

for x in range(x_start, x_end + 1):
    for y in range(y_start, y_end + 1):
        url = montar_url_tile(CAMADA, DATA_IMAGEM, "250m", ZOOM_LEVEL, y, x)
        print(f"Baixando {url}...")
        resp = requests.get(url)

        if resp.status_code == 200:
            img = Image.open(BytesIO(resp.content))
            caminho_tile = os.path.join(PASTA_IMAGENS, f"tile_{x}_{y}.png")
            img.save(caminho_tile)
            tiles_baixados.append((x, y, caminho_tile))
        else:
            print(f"Falha ao baixar tile {x},{y} ({resp.status_code})")

# 6. JUNTAR TILES EM UMA ÚNICA IMAGEM ----------------

if tiles_baixados:
    xs = [tile[0] for tile in tiles_baixados]
    ys = [tile[1] for tile in tiles_baixados]

    largura_total = (max(xs) - min(xs) + 1) * 256
    altura_total = (max(ys) - min(ys) + 1) * 256

    imagem_final = Image.new('RGB', (largura_total, altura_total))

    for x, y, caminho_tile in tiles_baixados:
        img = Image.open(caminho_tile)
        pos_x = (x - min(xs)) * 256
        pos_y = (y - min(ys)) * 256
        imagem_final.paste(img, (pos_x, pos_y))

    imagem_final.save("imagem_area_completa.png")
    print("Imagem final salva como 'imagem_area_completa.png'!")
else:
    print("Nenhum tile baixado.")
