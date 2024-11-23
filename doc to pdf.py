from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

def doc_to_pdf(doc_file, pdf_file):
    # Load the .docx file
    try:
        document = Document(doc_file)
    except Exception as e:
        print(f"Error loading document: {e}")
        return
    
    # Create a PDF file
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter
    
    # Read paragraphs from the .docx and write them to the PDF
    y = height - 40  # Start from the top margin
    line_spacing = 14  # Line spacing in PDF

    for paragraph in document.paragraphs:
        text = paragraph.text
        if text.strip():  # Skip empty lines
            if y < 40:  # Check if there's space for more text
                c.showPage()
                y = height - 40
            c.drawString(40, y, text[:1000])  # Limit text to fit the width
            y -= line_spacing

    c.save()
    print(f"Converted '{doc_file}' to '{pdf_file}' successfully.")

if __name__ == "__main__":
    input_file = input("Enter the path to the .docx file: ").strip()
    if not input_file.lower().endswith('.docx'):
        print("Please provide a valid .docx file.")
    elif not os.path.exists(input_file):
        print("File not found. Please check the path and try again.")
    else:
        output_file = input_file.replace('.docx', '.pdf')
        doc_to_pdf(input_file, output_file)
