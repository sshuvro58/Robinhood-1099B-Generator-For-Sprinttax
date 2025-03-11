import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os
import datetime
import re

def fill_pdf(template_pdf, data_dict, output_pdf):
    reader = PdfReader(template_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        writer.add_page(page)
        writer.update_page_form_field_values(writer.pages[i], data_dict)

    with open(output_pdf, 'wb') as f_out:
        writer.write(f_out)

def generate_pdfs_from_excel(excel_file, template_pdf, output_folder):
    df = pd.read_excel(excel_file)
    print(f"Loaded {len(df)} rows from Excel.")

    os.makedirs(output_folder, exist_ok=True)

    for index, row in df.iterrows():
        data_dict = {}

        data_dict['f2_1[0]'] = " Robinhood Markets Inc\n85 Willow Road\nMenlo Park, CA   94025"
        data_dict['f2_2[0]'] = "46-4364776"
        data_dict['f2_9[0]'] = "TX"
        data_dict['f2_16[0]'] = row["DESCRIPTION"]

        date_acquired = row["DATE ACQUIRED"]
        if not pd.isna(date_acquired):
            date_acquired = date_acquired.strftime('%m/%d/%Y') if hasattr(date_acquired, 'strftime') else str(date_acquired)
        else:
            date_acquired = datetime.date.today().strftime('%m/%d/%Y')

        data_dict['f2_17[0]'] = date_acquired
        data_dict['f2_18[0]'] = row["SALE DATE"]
        data_dict['f2_19[0]'] = row["COST BASIS"]
        data_dict['f2_20[0]'] = row["SALES PRICE"]
        data_dict['f2_23[0]'] = 0

        print(f"\nRow {index + 1}:")
        for k, v in data_dict.items():
            print(f"  {k} => {v}")

        safe_description = re.sub(r'[\\/*?:"<>|]', "_", str(row["DESCRIPTION"]))
        output_pdf = os.path.join(output_folder, f"{index+1}_filled_form_{safe_description}.pdf")

        fill_pdf(template_pdf, data_dict, output_pdf)

    print("\nDone! All PDFs created.")

# Example Usage
excel_file = 'robinhood.xlsx'
template_pdf = 'input_form.pdf'
output_folder = 'output_pdfs'

generate_pdfs_from_excel(excel_file, template_pdf, output_folder)
