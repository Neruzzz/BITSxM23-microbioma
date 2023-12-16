import pandas as pd
import openpyxl

# Reload the dataset
file_path = 'data/datasets/metabolite_microbe.xlsx'
data = pd.read_excel(file_path)

# Filter the columns
filtered_data = data[['compound_names', 'origin_type', 'genus', 'origin_species']]

# Save the filtered data to a new Excel file
output_file_path = 'data/outputs/filtered_metabolite_microbe.xlsx'
with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
    filtered_data.to_excel(writer, index=False)

# Open the saved Excel file and adjust column widths
wb = openpyxl.load_workbook(output_file_path)
ws = wb.active

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column letter

    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass

    adjusted_width = (max_length + 2)  # Adding a little extra space
    ws.column_dimensions[column].width = adjusted_width

wb.save(output_file_path)