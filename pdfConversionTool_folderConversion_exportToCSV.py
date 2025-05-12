from PyPDF2 import PdfReader, PdfWriter
from tkinter import *
import fitz  # PyMuPDF
import pandas as pd
import os

input_pdf_folder = r'c:\Users\INSB08203\OneDrive - WSP O365\Desktop\New folder'
output_pdf_folder = r'C:\Users\INSB08203\OneDrive - WSP O365\Desktop\New_Folder2'  # Corrected to use the input folder directly
rotation_angle = 180  # Rotate by 180 degrees clockwise

def is_scanned_page(page):
    text = page.get_text("text")
    return not bool(text.strip()) # Returns True if no text is found

def rotate_pdf_page(input_pdf_path, output_pdf_path, rotation_angle):

    total_scanned_pages = 0
    total_text_pages = 0

# Check if the page is scanned image
    pdf_document = fitz.open(input_pdf_path)
    print("Total pages in pdf doc: "+ str(len(pdf_document)))

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        if is_scanned_page(page):
            total_scanned_pages += 1
            # print(f"Page {page_num + 1} is likely scanned.")
        else:
            total_text_pages += 1
            # print(f"Page {page_num + 1} is not a scanned page.")
    pdf_document.close() 

    print(f"Total scanned pages: {total_scanned_pages}")
    print(f"Total text pages: {total_text_pages}")

    with open(input_pdf_path, 'rb') as input_file:
        reader = PdfReader(input_file)
        writer = PdfWriter()

        for i in range(len(reader.pages)):
            page = reader.pages[i]
            # Check if the page is scanned
            pdf_document = fitz.open(input_pdf_path)
            fitz_page = pdf_document[i]
            if is_scanned_page(fitz_page):
                page.rotate(rotation_angle)
            writer.add_page(page)
            pdf_document.close()

        with open(output_pdf_path, 'wb') as output_file:
            writer.write(output_file)
        print(f"Process complete. Rotated PDF saved to {output_pdf_path}")

    export_data = pd.DataFrame({ 'filename': [input_pdf_path],
                                'total_pages': [len(reader.pages)],
                                'total_scanned_pages': [total_scanned_pages],
                                'total_text_pages': [total_text_pages] })
    
    summary_csv_path = os.path.join(output_pdf_folder, 'summary.csv')
    if not os.path.exists(summary_csv_path):
        export_data.to_csv(summary_csv_path, index=False)  # Write header if file doesn't exist
    else:
        export_data.to_csv(summary_csv_path, mode='a', header=False, index=False)  # Append without header

# Iterate through all PDF files in the input folder
for files in os.listdir(input_pdf_folder):
    print(files)

    # Check if the file is a PDF
    # input_pdf_path = os.path.join(input_pdf_folder, files)
    if files.endswith('.pdf'):
        input_pdf_path = os.path.join(input_pdf_folder, files)  # Corrected path construction
        # output_pdf_path = input_pdf_path.replace('.pdf', '_rotated.pdf')
        output_pdf_path=output_pdf_folder + '\\' + files.replace('.pdf', '_rotated.pdf')
        rotate_pdf_page(input_pdf_path, output_pdf_path, rotation_angle)
            
