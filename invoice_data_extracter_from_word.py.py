import os
import re
from docx import Document
import pandas as pd
import docx2txt

def extract_invoice_details(doc_path):
    try:
        text = docx2txt.process(doc_path)
        
    except Exception as e:
        print(f"Error reading the document: {e}")

    invoice_number_pattern = r'INVOICE NUMBER\s*([\w\- ]+)\s'
    issue_date_pattern = r'ISSUE DATE\s*([\w, ]+)\s'

    invoice_number_match = re.search(invoice_number_pattern, text)
    issue_date_match = re.search(issue_date_pattern, text)

    invoice_number = invoice_number_match.group(1).strip() if invoice_number_match else 'Not found'
    issue_date = issue_date_match.group(1).strip() if issue_date_match else 'Not found'


    

    return invoice_number, issue_date

def process_directory(directory_path):
    data = { 'Invoice Number': [], 'Issue Date': []}
    errors = []

    for file_name in os.listdir(directory_path):
        if file_name.endswith('.docx'):
            file_path = os.path.join(directory_path, file_name)
            try:
                invoice_number, issue_date = extract_invoice_details(file_path)
                # data['File Name'].append(file_name)
                data['Invoice Number'].append(invoice_number)
                data['Issue Date'].append(issue_date)
            except Exception as e:
                errors.append((file_name, str(e)))
                print(f"Error processing {file_name}: {e}")

    df = pd.DataFrame(data)
    excel_file_path = os.path.join(directory_path, 'invoice_details.xlsx')
    df.to_excel(excel_file_path, index=False)
    print("Invoice details saved to Excel successfully!")

    if errors:
        error_log_path = os.path.join(directory_path, 'error_log.txt')
        with open(error_log_path, 'w') as error_log:
            for error in errors:
                error_log.write(f"File: {error[0]}, Error: {error[1]}\n")
        print(f"Errors encountered with some files. See {error_log_path} for details.")

# Path to the directory containing Word files
directory_path = r'E:\Project--reading a word files and extracting\Screaming-frog\Docx'
process_directory(directory_path)
