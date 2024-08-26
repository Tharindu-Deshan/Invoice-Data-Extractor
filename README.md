# Invoice Data Extraction Project

Overview

This project is designed to extract invoice details such as the invoice number and issue date from .docx files in a specified directory. The extracted data is then saved into an Excel file for further analysis. Additionally, any errors encountered during the extraction process are logged into a text file as well as we can convert PDF files in to docx files.

Requirements

Python 3.x

Required Python packages:

python-docx
docx2txt
pandas

pip install pdf2do(For PDF to docx conversion)


You can install the required packages using pip:
pip install python-docx docx2txt pandas

Project Structure

extract_invoice_details.py: The main script that processes the .docx files, extracts invoice details, and saves the data into an Excel file.
README.txt: This file contains information about the project and how to use it.

How to Use

1)Set the Directory Path: Update the directory_path variable in the script with the path to the directory containing your .docx files.

2)Run the Script: Execute the script to process the .docx files and extract invoice details.

python extract_invoice_details.py

Check the Output:

The extracted invoice data will be saved as invoice_details.xlsx in the specified directory.
Any errors encountered during processing will be logged in error_log.txt within the same directory.


Code Explanation

extract_invoice_details(doc_path):
Extracts invoice details (invoice number and issue date) from the .docx file specified by doc_path.
process_directory(directory_path):
Iterates through all .docx files in the specified directory.
Calls extract_invoice_details() for each file.
Saves the extracted data to an Excel file and logs any errors.

Notes
Ensure the invoice .docx files have the correct format for extraction.
The script uses regular expressions to identify and extract the required details.
PDF conversion is optional and requires additional setup
