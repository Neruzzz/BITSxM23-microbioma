import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
import numpy as np
from collections import defaultdict

# Load Data
excel_file_path = 'data/microbiota_trustable.xlsx'
info_df = pd.read_excel(excel_file_path, sheet_name='Sample info + Sperm quality')
sheet2_df = pd.read_excel(excel_file_path, sheet_name='Pylum-level microbiota')
sheet3_df = pd.read_excel(excel_file_path, sheet_name='Family-level microbiota')
sheet4_df = pd.read_excel(excel_file_path, sheet_name='Genus-level microbiota')

# Preprocess Data
def preprocess_data(df, info_df = False):
    df['Label'] = df['Sample ID'].apply(lambda x: 1 if x.startswith('CON') else 0)
    features = df.drop(['Sample ID', 'Label'], axis=1)
    # features.select_dtypes(['number'])
    
    if info_df:
        label_encoder = LabelEncoder()
        for column in features.columns:           
            # Check if the data type of the column is non-numeric (object type)
            if features[column].dtype == 'object' or pd.api.types.is_datetime64_any_dtype(df[column]):
                # Use label encoding to transform the categorical values to numerical values
                features[column] = label_encoder.fit_transform(features[column])
    
    labels = df['Label']
    return features, labels

sheet2_features, sheet2_labels = preprocess_data(sheet2_df)
sheet3_features, sheet3_labels = preprocess_data(sheet3_df)
sheet4_features, sheet4_labels = preprocess_data(sheet4_df)
info_features, info_labels = preprocess_data(info_df, True)

# # Train Random Forest Classifier
# def train_random_forest_classifier(features, labels):
#     clf = RandomForestClassifier(n_estimators=100, random_state=42)
#     clf.fit(features, labels)
#     return clf

# clf_info = train_random_forest_classifier(info_features, info_labels)
# clf_sheet2 = train_random_forest_classifier(sheet2_features, sheet2_labels)
# clf_sheet3 = train_random_forest_classifier(sheet3_features, sheet3_labels)
# clf_sheet4 = train_random_forest_classifier(sheet4_features, sheet4_labels)

# # Get Feature Importance
# def get_feature_importance(clf, features):
#     feature_importance = pd.DataFrame({'Feature': features.columns, 'Importance': clf.feature_importances_})
#     feature_importance = feature_importance.sort_values(by='Importance', ascending=False)
#     return feature_importance

# feature_importance_info = get_feature_importance(clf_info, info_features)
# feature_importance_sheet2 = get_feature_importance(clf_sheet2, sheet2_features)
# feature_importance_sheet3 = get_feature_importance(clf_sheet3, sheet3_features)
# feature_importance_sheet4 = get_feature_importance(clf_sheet4, sheet4_features)

# # Display Feature Importance
# print("Feature Importance - pylum level:")
# print(feature_importance_sheet2)

# print("\nFeature Importance - family level:")
# print(feature_importance_sheet3)

# print("\nFeature Importance - genus level:")
# print(feature_importance_sheet4)



from scipy.stats import spearmanr

# Calculate Spearman's rank correlation for each feature
def calculate_spearman_correlation(features, labels):
    correlations = {}
    for column in features.columns:
        correlation, _ = spearmanr(features[column], labels)
        correlations[column] = correlation
    return correlations

# Calculate correlation for each sheet
correlations_dict_info = calculate_spearman_correlation(info_features, info_labels)
correlations_dict_pylum = calculate_spearman_correlation(sheet2_features, sheet2_labels)
correlations_dict_family = calculate_spearman_correlation(sheet3_features, sheet3_labels)
correlations_dict_genus = calculate_spearman_correlation(sheet4_features, sheet4_labels)

# Sort features based on correlation
sorted_correlations_info  = sorted(correlations_dict_info.items(), key=lambda x: x[1], reverse=True)
sorted_correlations_pylum = sorted(correlations_dict_pylum.items(), key=lambda x: x[1], reverse=True)
sorted_correlations_family = sorted(correlations_dict_family.items(), key=lambda x: x[1], reverse=True)
sorted_correlations_genus = sorted(correlations_dict_genus.items(), key=lambda x: x[1], reverse=True)


# Display results

print("Spearman's Rank Correlation - info:")
# print(sorted_correlations_pylum)
for tup in sorted_correlations_info[:5]: print(tup[0], tup[1])
for tup in sorted_correlations_info[-5:]: print(tup[0], tup[1])

print("\nSpearman's Rank Correlation - Sheet 2:")
# print(sorted_correlations_pylum)
for tup in sorted_correlations_pylum[:5]: print(tup[0], tup[1])
for tup in sorted_correlations_pylum[-5:]: print(tup[0], tup[1])

print("\nSpearman's Rank Correlation - Sheet 3:")
for tup in sorted_correlations_family[:5]: print(tup[0], tup[1])
for tup in sorted_correlations_family[-5:]: print(tup[0], tup[1])
# print(sorted_correlations_family)

print("\nSpearman's Rank Correlation - Sheet 4:")
for tup in sorted_correlations_genus[:5]: print(tup[0], tup[1])
for tup in sorted_correlations_genus[-5:]: print(tup[0], tup[1])
# print(sorted_correlations_genus)



print(correlations_dict_family['Planococcaceae'])



microbe_file_path = 'C:\Diego\LLN\ERASMUS\Cours\SIRI\lamarato\seminal\git\BITSxM23-microbioma\data\datasets\metabolite_microbe.xlsx'

microbe_df = pd.read_excel(microbe_file_path)

    
cluster_sets_fertile = []
clusters_score_dict = defaultdict(lambda: [0,set()])
genus_clusters = defaultdict(lambda:set())
for genus,corr in sorted_correlations_genus:
    count_occurrences = (microbe_df['genus'] == genus).sum()

    if count_occurrences>0:
        # print(genus,corr,count_occurrences)
        filtered_rows = microbe_df.loc[microbe_df['genus'] == genus]
        clusters = set(filtered_rows['compound_node_id'].tolist())
        cluster_sets_fertile.append((genus,corr,clusters))
        
        for cluster in clusters:
            clusters_score_dict[cluster][0] += corr
            clusters_score_dict[cluster][1].add(genus)
            genus_clusters[cluster].add(genus)
    

# print(genus_clusters)
# print(cluster_sets_fertile)
scores = sorted(clusters_score_dict.items(), key=lambda x: x[1][0], reverse=True)
print(scores[:5])



