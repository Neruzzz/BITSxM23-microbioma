import pandas as pd
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage

# Load the dataset
file_path = 'data/datasets/microbiota_trustable.xlsx'  # Adjust the file path as needed
sheet_names = ['Pylum-level microbiota', 'Family-level microbiota', 'Genus-level microbiota']

# Process each sheet
for sheet in sheet_names:
    # Load the data
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Extract the fertility status as labels
    fertility_labels = df.iloc[:, 0].astype(str)

    # Drop the first two columns to focus only on microbiota data
    microbiota_data = df.drop(columns=[df.columns[0], df.columns[1]])

    # Perform hierarchical clustering using only the numerical data
    Z = linkage(microbiota_data, 'ward')

    # Plot the dendrogram
    plt.figure(figsize=(10, 7))
    plt.title(f'Dendrogram for {sheet}')
    dendrogram(Z, labels=fertility_labels.tolist(), leaf_rotation=90, leaf_font_size=10)
    plt.yscale('log')
    plt.ylim(0.5, 500)
    plt.xlabel('Subjects (Fertility Status)')
    plt.ylabel('Distance')

    # Save the plot to a file
    plt.savefig(f'data/outputs/graphs/dendrograms/{sheet.replace(" ", "-")}.png')
    plt.close()  # Close the plot to avoid overlapping of plots
