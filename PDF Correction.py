import os
import fitz  # PyMuPDF
import PyPDF2
import pandas as pd

folder_path = input("Please input the woking path: ")
#r'C:\Pitabasa Data\Python\Test'  # Update this if needed

# Get list of PDF files in the folder
PDFs = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

# Initialize report data
report_data = []

# --- Function to check page rotation using PyPDF2 ---
def get_non_180_pages(pdf_path):
    non_180_pages = []
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for i, page in enumerate(reader.pages):
            rotation = page.get('/Rotate', 0)
            if rotation != 180:
                non_180_pages.append(i + 1)  # Page numbers are 1-based
    return non_180_pages, len(reader.pages)

# --- Function to find blank pages using PyMuPDF ---
def get_blank_pages(pdf_path):
    blank_pages = []
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        page = doc.load_page(i)
        text = page.get_text("text")
        images = page.get_images(full=True)
        if not text.strip() and not images:
            blank_pages.append(i + 1)  # 1-based index
    doc.close()
    return blank_pages

# --- Function to auto-rotate and overwrite PDFs ---
def auto_rotate_pdf(input_pdf_path, output_pdf_path, non_180_pages):
    doc = fitz.open(input_pdf_path)
    for page_num in range(len(doc)):
        if (page_num + 1) in non_180_pages:
            page = doc.load_page(page_num)
            rotation = page.rotation
            print(f"In {os.path.basename(input_pdf_path)}: Page {page_num + 1} rotated by {rotation}°, correcting to 180°.")
            page.set_rotation(180)
    doc.save(output_pdf_path)
    doc.close()

# --- Process each PDF file ---
for pdf in PDFs:
    pdf_path = os.path.join(folder_path, pdf)
    temp_path = os.path.join(folder_path, f"Modified_{pdf}")

    # Analyze pages
    non_180_pages, total_pages = get_non_180_pages(pdf_path)
    blank_pages = get_blank_pages(pdf_path)

    # Rotate pages if needed
    if non_180_pages:
        auto_rotate_pdf(pdf_path, temp_path, non_180_pages)
        os.remove(pdf_path)
        os.rename(temp_path, pdf_path)

    # Append data for reporting
    report_data.append({
        'PDF Name': pdf,
        'Total Pages': total_pages,
        'Non-180 Page Count': len(non_180_pages),
        'Non-180 Page Numbers': ', '.join(map(str, non_180_pages)) if non_180_pages else 'None',
        'Blank Page Numbers': ', '.join(map(str, blank_pages)) if blank_pages else 'None'
    })

# --- Create DataFrame and write to Excel ---
df = pd.DataFrame(report_data)
report_path = os.path.join(folder_path, 'PDF_Detailed_Report.xlsx')
df.to_excel(report_path, index=False)

print(f"Excel report generated at: {report_path}")
