from fillpdf import fillpdfs
import os

input_pdf_path = 'Source.pdf'
output_directory = 'output_pdfs'

if not os.path.exists(output_directory):
    os.makedirs(output_directory)

form_fields = fillpdfs.get_form_fields(input_pdf_path)
print("PDF Form Fields: ", form_fields)

data_dict = {}

# Loop over form fields and assign sequential numbering
for field_num, field_name in enumerate(form_fields.keys(), start=1):
    data_dict[field_name] = str(field_num)  # Assign the number as a string

# Define the output PDF path
output_pdf_path = os.path.join(output_directory, 'filled_form.pdf')

# Fill the PDF with the data_dict (numbers)
fillpdfs.write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict)

print(f'Generated {output_pdf_path}')
print("PDF filled successfully with numbers!")
