import pandas as pd
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.metrics import balanced_accuracy_score, accuracy_score
import matplotlib.pyplot as plt

# Load Data
excel_file_path = 'data/datasets/microbiota_trustable.xlsx'
sheet2_df = pd.read_excel(excel_file_path, sheet_name='Pylum-level microbiota')
sheet3_df = pd.read_excel(excel_file_path, sheet_name='Family-level microbiota')
sheet4_df = pd.read_excel(excel_file_path, sheet_name='Genus-level microbiota')

# Preprocess Data
def preprocess_data(df):
    df['Label'] = df['Sample ID'].apply(lambda x: 1 if x.startswith('CON') else 0)
    features = df.drop(['Sample ID', 'Label'], axis=1)
    labels = df['Label']
    return features, labels

sheet2_features, sheet2_labels = preprocess_data(sheet2_df)
sheet3_features, sheet3_labels = preprocess_data(sheet3_df)
sheet4_features, sheet4_labels = preprocess_data(sheet4_df)

# Perform PCA
def perform_pca(features, n_components=2):
    pca = PCA(n_components=n_components)
    features_pca = pca.fit_transform(features)
    return features_pca, pca

# Apply PCA for each sheet
sheet2_features_pca, pca_sheet2 = perform_pca(sheet2_features)
sheet3_features_pca, pca_sheet3 = perform_pca(sheet3_features)
sheet4_features_pca, pca_sheet4 = perform_pca(sheet4_features)

# Perform KMeans clustering
def perform_kmeans(features, n_clusters=2):
    kmeans = KMeans(n_clusters=n_clusters, random_state=42)
    labels = kmeans.fit_predict(features)
    return labels

# Apply KMeans clustering for each sheet
sheet2_labels_kmeans = perform_kmeans(sheet2_features_pca)
sheet3_labels_kmeans = perform_kmeans(sheet3_features_pca)
sheet4_labels_kmeans = perform_kmeans(sheet4_features_pca)

def visualize_clusters(features_pca, labels_kmeans, true_labels, sheet_name):
    plt.figure(figsize=(8, 6.5))

    # Plot true labels with different marker styles
    plt.scatter(features_pca[true_labels == 0, 0], features_pca[true_labels == 0, 1], c='blue', label='True Label 0', marker='o', s=100, alpha=0.7)
    plt.scatter(features_pca[true_labels == 1, 0], features_pca[true_labels == 1, 1], c='orange', label='True Label 1', marker='s', s=100, alpha=0.7)

    # Plot KMeans labels with different colors
    plt.scatter(features_pca[labels_kmeans == 0, 0], features_pca[labels_kmeans == 0, 1], c='red', label='KMeans Cluster 0', marker='^', s=100, alpha=0.7)
    plt.scatter(features_pca[labels_kmeans == 1, 0], features_pca[labels_kmeans == 1, 1], c='green', label='KMeans Cluster 1', marker='v', s=100, alpha=0.7)

    plt.title(f'Clustering Comparison - {sheet_name}')
    plt.xlabel('Principal Component 1')
    plt.ylabel('Principal Component 2')

    # Evaluate KMeans accuracy
    accuracy = accuracy_score(true_labels, labels_kmeans)
    note = f'KMeans Accuracy: {accuracy:.2f}'
    plt.text(0.9, -0.1, note, fontsize=10, ha='center', transform=plt.gca().transAxes)

    plt.legend(loc='upper right')
    plt.show()

# Visualize clusters for each sheet
visualize_clusters(sheet2_features_pca, sheet2_labels_kmeans, sheet2_labels, 'Sheet 2')
visualize_clusters(sheet3_features_pca, sheet3_labels_kmeans, sheet3_labels, 'Sheet 3')
visualize_clusters(sheet4_features_pca, sheet4_labels_kmeans, sheet4_labels, 'Sheet 4')
