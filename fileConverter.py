import os
from PIL import Image
import pandas as pd
from fpdf import FPDF
from PyPDF2 import PdfReader
from docx import Document
from docx2pdf import convert as docx_to_pdf
from pdf2docx import Converter as pdf_to_docx

def convert_file(input_file, output_file):
    # Get file extensions
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    # Image 
    if input_ext in ['.jpg', '.jpeg', '.png'] and output_ext in ['.jpg', '.jpeg', '.png']:
        convert_image(input_file, output_file)
    # Text to PDF
    elif input_ext == '.txt' and output_ext == '.pdf':
        convert_text_to_pdf(input_file, output_file)
    # PDF to Text
    elif input_ext == '.pdf' and output_ext == '.txt':
        convert_pdf_to_text(input_file, output_file)
    # CSV to Excel
    elif input_ext == '.csv' and output_ext == '.xlsx':
        convert_csv_to_excel(input_file, output_file)
    # Excel to CSV
    elif input_ext == '.xlsx' and output_ext == '.csv':
        convert_excel_to_csv(input_file, output_file)
    # Word to PDF
    elif input_ext == '.docx' and output_ext == '.pdf':
        convert_docx_to_pdf(input_file, output_file)
    # PDF to Word
    elif input_ext == '.pdf' and output_ext == '.docx':
        convert_pdf_to_docx(input_file, output_file)
    else:
        print(f"Conversion from {input_ext} to {output_ext} is not supported yet.")

def convert_image(input_file, output_file):
    image = Image.open(input_file)
    image.save(output_file)
    print(f"Converted {input_file} to {output_file}")

def convert_text_to_pdf(input_file, output_file):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    with open(input_file, 'r') as file:
        for line in file:
            pdf.cell(200, 10, txt=line, ln=True)
    pdf.output(output_file)
    print(f"Converted {input_file} to {output_file}")

def convert_pdf_to_text(input_file, output_file):
    reader = PdfReader(input_file)
    with open(output_file, 'w') as file:
        for page in reader.pages:
            file.write(page.extract_text())
    print(f"Converted {input_file} to {output_file}")

def convert_csv_to_excel(input_file, output_file):
    df = pd.read_csv(input_file)
    df.to_excel(output_file, index=False)
    print(f"Converted {input_file} to {output_file}")

def convert_excel_to_csv(input_file, output_file):
    df = pd.read_excel(input_file)
    df.to_csv(output_file, index=False)
    print(f"Converted {input_file} to {output_file}")

def convert_docx_to_pdf(input_file, output_file):
    docx_to_pdf(input_file, output_file)
    print(f"Converted {input_file} to {output_file}")

def convert_pdf_to_docx(input_file, output_file):
    pdf_to_word = pdf_to_docx(input_file)
    pdf_to_word.convert(output_file)
    pdf_to_word.close()
    print(f"Converted {input_file} to {output_file}")

def main():
    print("Choose the type of conversion:")
    print("1. Image conversion (.jpg, .jpeg, .png)")
    print("2. Text to PDF (.txt to .pdf)")
    print("3. PDF to Text (.pdf to .txt)")
    print("4. CSV to Excel (.csv to .xlsx)")
    print("5. Excel to CSV (.xlsx to .csv)")
    print("6. Microsoft Word to PDF (.docx to .pdf)")
    print("7. PDF to Microsoft Word (.pdf to .docx)")

    choice = input("Enter the number of your choice: ")

    if choice == '1':
        input_file = input("Enter the path of the image file to be converted: ")
        output_file = input("Enter the output file path with the desired extension (.jpg, .jpeg, .png): ")
    elif choice == '2':
        input_file = input("Enter the path of the text file to be converted: ")
        output_file = input("Enter the output file path with the .pdf extension: ")
    elif choice == '3':
        input_file = input("Enter the path of the PDF file to be converted: ")
        output_file = input("Enter the output file path with the .txt extension: ")
    elif choice == '4':
        input_file = input("Enter the path of the CSV file to be converted: ")
        output_file = input("Enter the output file path with the .xlsx extension: ")
    elif choice == '5':
        input_file = input("Enter the path of the Excel file to be converted: ")
        output_file = input("Enter the output file path with the .csv extension: ")
    elif choice == '6':
        input_file = input("Enter the path of the Microsoft Word file to be converted: ")
        output_file = input("Enter the output file path with the .pdf extension: ")
    elif choice == '7':
        input_file = input("Enter the path of the PDF file to be converted: ")
        output_file = input("Enter the output file path with the .docx extension: ")
    else:
        print("Invalid choice. Please restart the program and try again.")
        return

    convert_file(input_file, output_file)

if __name__ == "__main__":
    main()
