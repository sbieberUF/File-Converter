import os
from PIL import Image
import pandas as pd
from fpdf import FPDF
from pdfminer.high_level import extract_text
from docx2pdf import convert
from pdf2docx import Converter
import csv

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
    # CSV to SQL
    elif input_ext == '.csv' and output_ext == '.sql':
        convert_csv_to_sql(input_file, output_file)
    # PDF to Word
    elif input_ext == '.pdf' and output_ext == '.docx':
        convert_pdf_to_word(input_file, output_file)
    # Word to PDF
    elif input_ext == '.docx' and output_ext == '.pdf':
        convert_word_to_pdf(input_file, output_file)
    else:
        print(f"Conversion from {input_ext} to {output_ext} is not supported yet.")

def convert_image(input_file, output_file):
    image = Image.open(input_file)
    
    # Check if the image has an alpha channel (transparency) and convert to RGB
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    
    image.save(output_file)
    print(f"Converted {input_file} to {output_file}")

def convert_text_to_pdf(input_file, output_file):
    pdf = FPDF()
    pdf.add_page()

    # Unicode-compatible font
    pdf.add_font('DejaVuSans', '', 'DejaVuSans.ttf')
    pdf.set_font('DejaVuSans', size=12)

    try:
        # Open the input text file with UTF-8 encoding
        with open(input_file, 'r', encoding='utf-8') as file:
            for line in file:
                pdf.multi_cell(0, 10, text=line)  # Use multi_cell to handle line breaks and longer lines
        pdf.output(output_file)
        print(f"Converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Failed to convert {input_file} to PDF: {e}")


def convert_pdf_to_text(input_file, output_file):
    try:
        text = extract_text(input_file)
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(text)
        print(f"Converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Failed to convert {input_file} to text: {e}")

def convert_csv_to_excel(input_file, output_file):
    df = pd.read_csv(input_file)
    df.to_excel(output_file, index=False)
    print(f"Converted {input_file} to {output_file}")

def convert_excel_to_csv(input_file, output_file):
    df = pd.read_excel(input_file)
    df.to_csv(output_file, index=False)
    print(f"Converted {input_file} to {output_file}")


def convert_csv_to_sql(input_file, output_file):
    table_name = input("Enter the SQL table name: ")
    
    with open(input_file, mode='r') as file, open(output_file, mode='w') as sql_file:
        csv_reader = csv.reader(file)
        headers = next(csv_reader)  
        for row in csv_reader:
            # Escape single quotes in the values
            row = [value.replace("'", "") for value in row]
            # Create the insert statement
            insert_statement = f"INSERT INTO {table_name} ({', '.join(headers)}) VALUES ('" + "', '".join(row) + "');\n"
            # Write the insert statement to the SQL file
            sql_file.write(insert_statement)
    print(f"Converted {input_file} to {output_file} with SQL insert statements")

def convert_pdf_to_word(input_file, output_file):
    try:
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        print(f"Converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Failed to convert {input_file} to DOCX: {e}")    

def convert_word_to_pdf(input_file, output_file):
    try:
        # Convert the Word file to PDF
        convert(input_file, output_file)
        print(f"Converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Failed to convert {input_file} to PDF: {e}")    

def main():
    print("Choose the type of conversion:\n")
    print("1. Image conversion (.jpg, .jpeg, .png)")
    print("2. Text to PDF (.txt to .pdf)")
    print("3. PDF to Text (.pdf to .txt)")
    print("4. CSV to Excel (.csv to .xlsx)")
    print("5. Excel to CSV (.xlsx to .csv)")
    print("6. CSV to SQL Insert Statements (.csv to .sql)")
    print("7. PDF to Microsoft Word (.pdf to .docx)")
    print("8. Microsoft Word to PDF (.docx to .pdf)")

    choice = input("Enter the number of your choice: ")

    if choice == '1':
        input_file = input("Enter the path of the image file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the desired extension (.jpg, .jpeg, .png): ").replace('"', '')
    elif choice == '2':
        input_file = input("Enter the path of the text file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .pdf extension: ").replace('"', '')
    elif choice == '3':
        input_file = input("Enter the path of the PDF file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .txt extension: ").replace('"', '')
    elif choice == '4':
        input_file = input("Enter the path of the CSV file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .xlsx extension: ").replace('"', '')
    elif choice == '5':
        input_file = input("Enter the path of the Excel file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .csv extension: ").replace('"', '')
    elif choice == '6':
        input_file = input("Enter the path of the CSV file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .sql extension: ").replace('"', '')
    elif choice == '7':
        input_file = input("Enter the path of the PDF file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .docx extension: ").replace('"', '')
    elif choice == '8':
        input_file = input("Enter the path of the Microsoft Word file to be converted: ").replace('"', '')
        output_file = input("Enter the output file path with the .pdf extension: ").replace('"', '')
    else:
        print("Invalid choice. Please restart the program and try again.")
        return

    convert_file(input_file, output_file)

if __name__ == "__main__":
    main()
