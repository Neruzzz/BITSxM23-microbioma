import pandas as pd
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.metrics import accuracy_score, balanced_accuracy_score, precision_score
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

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

def perform_pca(features, n_components=3):
    pca = PCA(n_components=n_components)
    features_pca = pca.fit_transform(features)
    return features_pca, pca

def perform_kmeans(features, n_clusters=2):
    kmeans = KMeans(n_clusters=n_clusters, random_state=42)
    labels = kmeans.fit_predict(features)
    return labels

def visualize_clusters(features_pca, labels_kmeans, true_labels, sheet_name):
    fig = plt.figure(figsize=(10, 8))
    ax = fig.add_subplot(111, projection='3d')

    # Plot true labels
    ax.scatter(features_pca[true_labels == 0, 0], features_pca[true_labels == 0, 1], features_pca[true_labels == 0, 2],
               c='blue', label='True Label 0', marker='o', s=100, alpha=0.7)
    ax.scatter(features_pca[true_labels == 1, 0], features_pca[true_labels == 1, 1], features_pca[true_labels == 1, 2],
               c='orange', label='True Label 1', marker='s', s=100, alpha=0.7)

    # Plot KMeans labels
    ax.scatter(features_pca[labels_kmeans == 0, 0], features_pca[labels_kmeans == 0, 1], features_pca[labels_kmeans == 0, 2],
               c='red', label='KMeans Cluster 0', marker='^', s=100, alpha=0.7)
    ax.scatter(features_pca[labels_kmeans == 1, 0], features_pca[labels_kmeans == 1, 1], features_pca[labels_kmeans == 1, 2],
               c='green', label='KMeans Cluster 1', marker='v', s=100, alpha=0.7)

    ax.set_title(f'Clustering Comparison - {sheet_name}')
    ax.set_xlabel('Principal Component 1')
    ax.set_ylabel('Principal Component 2')
    ax.set_zlabel('Principal Component 3')

    # Convert KMeans cluster assignments to binary labels (0 or 1)
    predicted_labels = (labels_kmeans == labels_kmeans.max()).astype(int)

    # Evaluate regular classifier accuracy
    classifier_accuracy = accuracy_score(true_labels, predicted_labels)
    note = f'Classifier Accuracy: {classifier_accuracy:.2f}'
    # ax.text2D(0.5, -0.1, note, fontsize=10, ha='center', transform=ax.transAxes)

    ax.legend(loc='upper right')
    plt.show()

# Apply PCA and KMeans clustering for each sheet
sheet2_features, sheet2_labels = preprocess_data(sheet2_df)
sheet2_features_pca, pca_sheet2 = perform_pca(sheet2_features)
sheet2_labels_kmeans = perform_kmeans(sheet2_features_pca)
classifier_accuracy = balanced_accuracy_score(sheet2_labels, sheet2_labels_kmeans)
print(f'Cluster Accuracy for Pylum-level microbiota: {classifier_accuracy:.2f}')

sheet3_features, sheet3_labels = preprocess_data(sheet3_df)
sheet3_features_pca, pca_sheet3 = perform_pca(sheet3_features)
sheet3_labels_kmeans = perform_kmeans(sheet3_features_pca)
classifier_accuracy = balanced_accuracy_score(sheet3_labels, sheet3_labels_kmeans)
print(f'Cluster Accuracy for Family-level microbiota: {classifier_accuracy:.2f}')

sheet4_features, sheet4_labels = preprocess_data(sheet4_df)
sheet4_features_pca, pca_sheet4 = perform_pca(sheet4_features)
sheet4_labels_kmeans = perform_kmeans(sheet4_features_pca)
classifier_accuracy = balanced_accuracy_score(sheet4_labels, sheet4_labels_kmeans)
print(f'Cluster Accuracy for Genus-level microbiota: {classifier_accuracy:.2f}')

# Visualize clusters for each sheet in three dimensions
visualize_clusters(sheet2_features_pca, sheet2_labels_kmeans, sheet2_labels, 'Pylum-level')
visualize_clusters(sheet3_features_pca, sheet3_labels_kmeans, sheet3_labels, 'Family-level')
visualize_clusters(sheet4_features_pca, sheet4_labels_kmeans, sheet4_labels, 'Genus-level')
