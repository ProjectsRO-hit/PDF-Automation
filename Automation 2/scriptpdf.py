import pandas as pd
from fillpdf import fillpdfs
import os

# Load the Excel data
df = pd.read_excel('form_fields.xlsx')

# Path to the source PDF and output directory
input_pdf_path = 'Source.pdf'
output_directory = 'output_pdfs'

# Create output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Read the list of text box names from the file (assuming one name per line)
with open('field_names.txt', 'r') as file:
    field_names = [line.strip() for line in file]

# Ensure there are enough columns in the DataFrame
num_columns = len(df.columns)

# Ensure that the number of field names matches the number of data columns in the DataFrame
if len(field_names) < num_columns:
    raise ValueError("Not enough field names for the columns in the DataFrame.")

# Loop through each row in the Excel file
for index, row in df.iterrows():
    data_dict = {}
    
    # Loop through each field name and map it to the corresponding DataFrame column
    for i, field_name in enumerate(field_names):
        column_name = f'a{i+1}'  # Assume DataFrame columns are named 'a1', 'a2', etc.
        if column_name in df.columns:
            data_dict[field_name] = row[column_name]
    
    # Path for the output PDF
    output_pdf_path = os.path.join(output_directory, f'filled_form_{index + 1}.pdf')

    # Fill the PDF with data
    fillpdfs.write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict)
    print(f'Generated {output_pdf_path}')

print("All PDFs generated successfully!")
