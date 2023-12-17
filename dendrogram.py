import pandas as pd
from sklearn.metrics import accuracy_score
from scipy.cluster.hierarchy import linkage, dendrogram, cut_tree
import matplotlib.pyplot as plt

# Load Data
excel_file_path = 'data/microbiota_trustable.xlsx'
sheet2_df = pd.read_excel(excel_file_path, sheet_name='Pylum-level microbiota')
sheet3_df = pd.read_excel(excel_file_path, sheet_name='Family-level microbiota')
sheet4_df = pd.read_excel(excel_file_path, sheet_name='Genus-level microbiota')

# Preprocess Data
def preprocess_data(df):
    df['Label'] = df['Sample ID'].apply(lambda x: 1 if x.startswith('CON') else 0)
    features = df.drop(['Sample ID', 'Label'], axis=1)
    labels = df['Label']
    return features, labels

def hierarchical_clustering(features, true_labels, sheet_name):
    # Calculate linkage matrix
    linkage_matrix = linkage(features, method='ward')

    # Plot dendrogram
    plt.figure(figsize=(10, 8))
    dendrogram(linkage_matrix, labels=true_labels.values, orientation='top', color_threshold=0, above_threshold_color='grey')

    plt.title(f'Hierarchical Clustering Dendrogram - {sheet_name}')
    plt.xlabel('Sample ID')
    plt.ylabel('Distance')

    plt.ylim(0.2, 200)
    plt.yscale('log')

    # Calculate regular classifier accuracy
    predicted_labels = cut_tree(linkage_matrix, n_clusters=2).flatten()
    classifier_accuracy = accuracy_score(true_labels, predicted_labels)
    note = f'Classifier Accuracy: {classifier_accuracy:.2f}'
    plt.text(0.5, -0.1, note, fontsize=10, ha='center', transform=plt.gca().transAxes)

    plt.show()

# Apply hierarchical clustering and visualize dendrogram for each sheet
sheet2_features, sheet2_labels = preprocess_data(sheet2_df)
hierarchical_clustering(sheet2_features, sheet2_labels, 'Sheet 2')

sheet3_features, sheet3_labels = preprocess_data(sheet3_df)
hierarchical_clustering(sheet3_features, sheet3_labels, 'Sheet 3')

sheet4_features, sheet4_labels = preprocess_data(sheet4_df)
hierarchical_clustering(sheet4_features, sheet4_labels, 'Sheet 4')
