import os
from PIL import Image
import pandas as pd
from fpdf import FPDF 

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
    # CSV to Excel
    elif input_ext == '.csv' and output_ext == '.xlsx':
        convert_csv_to_excel(input_file, output_file)
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

def convert_csv_to_excel(input_file, output_file):
    df = pd.read_csv(input_file)
    df.to_excel(output_file, index=False)
    print(f"Converted {input_file} to {output_file}")

def main():
    print("Choose the type of conversion:")
    print("1. Image conversion (.jpg, .jpeg, .png)")
    print("2. Text to PDF (.txt to .pdf)")
    print("3. CSV to Excel (.csv to .xlsx)")

    choice = input("Enter the number of your choice: ")

    if choice == '1':
        input_file = input("Enter the path of the image file to be converted: ")
        output_file = input("Enter the output file path with the desired extension (.jpg, .jpeg, .png): ")
    elif choice == '2':
        input_file = input("Enter the path of the text file to be converted: ")
        output_file = input("Enter the output file path with the .pdf extension: ")
    elif choice == '3':
        input_file = input("Enter the path of the CSV file to be converted: ")
        output_file = input("Enter the output file path with the .xlsx extension: ")
    else:
        print("Invalid choice. Please restart the program and try again.")
        return

    convert_file(input_file, output_file)

if __name__ == "__main__":
    main()
