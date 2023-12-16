import pandas as pd
from scipy.stats import spearmanr
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Load the dataset and prepare to analyze other sheets
file_path = 'data/datasets/microbiota_trustable.xlsx'
sheets = ['Pylum-level microbiota', 'Family-level microbiota', 'Genus-level microbiota']
correlations_data = {"Pylum": [], "Family": [], "Genus": []}

# Function to calculate correlations
def calculate_correlations(data):
    fertility_status = data['Fertility']
    correlations = {}
    for taxon in data.columns[2:]:  # Start processing data from the third column
        correlation, _ = spearmanr(data[taxon], fertility_status)
        correlations[taxon] = correlation

    # Sort taxa by their correlation coefficient from most positive to most negative
    sorted_taxa = sorted(correlations.items(), key=lambda x: x[1], reverse=True)
    return sorted_taxa

# Process each sheet
for sheet in sheets:
    # Load each sheet
    data = pd.read_excel(file_path, sheet_name=sheet)

    # Calculate correlations
    top_taxa = calculate_correlations(data)
    level = sheet.split('-')[0]  # Extract level (Pylum, Family, Genus)
    correlations_data[level].extend(top_taxa)

# Function to save the correlations in an Excel file
def save_correlations_to_excel(correlations, filename):
    with pd.ExcelWriter(filename) as writer:
        for level, data in correlations.items():
            df = pd.DataFrame(data, columns=[f'{level} Taxon', 'Spearman Correlation'])
            df.to_excel(writer, sheet_name=level, index=False)

# Save the correlations to Excel files
save_correlations_to_excel(correlations_data, 'data/outputs/microbiota_correlations.xlsx')
