import pandas as pd
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage







# Función para crear dendrogramas
def create_dendrogram(data, level, fertility_status, status_label):
    # Filtrar datos basados en el estado de fertilidad y eliminar la segunda columna
    filtered_data = data[data.iloc[:, 0] == fertility_status].iloc[:, 2:]
    print(filtered_data)
    # Clustering jerárquico
    linked = linkage(filtered_data.T, method='ward')  # Transponer datos para agrupar por taxonomia

    # Dendrograma
    plt.figure(figsize=(15, 10))
    dendrogram(linked, orientation='top', labels=filtered_data.columns, distance_sort='descending', show_leaf_counts=True)
    plt.title(f'Dendrograma de clustering nivel de {level} - {status_label}')
    plt.xlabel('Tipo de ' + level)
    plt.ylabel('Distancia de Ward (log)')
    plt.yscale('log')
    plt.ylim(0.0001,1000)
    plt.xticks(rotation=90)
    plt.savefig(f'data/outputs/graphs/dendograms/dendrograma_{level}_{status_label}.png')
    plt.close()

# Cargar el dataset
file_path = 'data/datasets/microbiota_trustable.xlsx'
sheets = ['Pylum-level microbiota', 'Family-level microbiota', 'Genus-level microbiota']

for sheet in sheets:
    level = sheet.split('-')[0]  # Extraer el nivel (Pylum, Family, Genus)
    data = pd.read_excel(file_path, sheet_name=sheet)

    # Crear dendrogramas para muestras fértiles e infértiles
    create_dendrogram(data, level, 1, 'Fertile')
    create_dendrogram(data, level, 0, 'Infertile')
