import os
from pdf2docx import Converter

def convert_pdf_to_word(pdf_folder, word_folder):
    # Check if the output folder exists, if not, create it
    if not os.path.exists(word_folder):
        os.makedirs(word_folder)

    # List all PDF files in the pdf_folder
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
    
    # Process each file
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        word_path = os.path.join(word_folder, pdf_file.replace('.pdf', '.docx'))
        
        # Create a Converter object and convert the PDF to Word
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        
        print(f"Converted {pdf_file} to {word_path}")

# Specify the folder containing PDFs and the folder to save Word documents
pdf_folder = r'E:\Project--reading a word files and extracting\pdf to word\pdf'
word_folder = r'E:\Project--reading a word files and extracting\pdf to word\word'

convert_pdf_to_word(pdf_folder, word_folder)
