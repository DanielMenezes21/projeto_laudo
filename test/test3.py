import geopandas as gpd
import matplotlib.pyplot as plt
from geopy import Point

# Leitura dos arquivos
propriedade = gpd.read_file("C:\\Users\\DESKTOP\\Desktop\\automacao_laudo\\anexos\\Processo nº 123456789 - JOELSON SOUSA JUNIOR\\CERT_INTEIRO_TEOR_M.11173.shp")
solos = gpd.read_file("test\\solo_tocantins\\solos_to_1000.json")

# Define CRS da propriedade para Web Mercator, depois converte para WGS84
propriedade = propriedade.set_crs(epsg=3857, allow_override=True)
propriedade = propriedade.to_crs(epsg=4326)

print(f"propriedade CRS: {propriedade.crs}")
print(f"solos CRS: {solos.crs}")

# Garante que ambos estejam no mesmo sistema de referência
if propriedade.crs != solos.crs:
    propriedade = propriedade.to_crs(solos.crs)
    print(f"Propriedade reprojetada para: {propriedade.crs}")
else:
    print("CRS já são compatíveis")

# Interseção espacial
intersecao = gpd.sjoin(propriedade, solos, how="inner", predicate="intersects")

print("Tipos de solo encontrados na propriedade:")
print(intersecao['descricao'].unique())

# Salva interseção em arquivo GeoJSON
intersecao.to_file("solos_na_propriedade.geojson", driver='GeoJSON')
# -------------------------------
# Extração e exibição das coordenadas (centroides) em graus decimais
# -------------------------------
# Converte para WGS84 se necessário
propriedade = propriedade.to_crs(epsg=4326)

propriedade_proj = propriedade.to_crs(epsg=31982)

# Calcular centroides com precisão
centroides = propriedade_proj.geometry.centroid

# Voltar para WGS84 para imprimir coordenadas em graus decimais
centroides = gpd.GeoSeries(centroides, crs=31982).to_crs(epsg=4326)

print("\nCoordenadas centrais (graus decimais):")
for i, centroide in enumerate(centroides):
    ponto = Point(centroide.y, centroide.x)  # latitude, longitude
    print(f"Coordenada {i+1}: Latitude {round(ponto.latitude, 6)}, Longitude {round(ponto.longitude, 6)}")

base = solos.plot(color='lightgray')
propriedade.plot(ax=base, color='red')
plt.show()
