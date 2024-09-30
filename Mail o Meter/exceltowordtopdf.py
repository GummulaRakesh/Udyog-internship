import pandas as pd
from docx import Document
from docx.shared import Pt
import comtypes.client


# Part 1: Excel Data Reading
def read_excel(file_path, sheet_name):
    print(f"Reading Excel file: {file_path}")
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Handle missing values: fill NaNs with a default value
    df.fillna('Missing', inplace=True)

    # Check for duplicate rows
    if df.duplicated().any():
        print("Warning: There are duplicate rows in the dataset.")

    # Check for empty columns
    if df.isnull().values.any():
        print("Warning: There are missing values in the dataset.")

    return df

# Part 2: Word Document Generation
def generate_word_report(excel_data, output_docx, output_pdf):
    # Create a new Word document
    doc = Document()

    # Set a title for the document
    doc.add_heading('Excel Data Report\n', level=1)

    # Iterate through Excel data and add content to the Word document
    for i, row in excel_data.iterrows():
        # Add a paragraph for each row of data
        doc.add_paragraph(f"Name: {row['Name']}, \nAge: {row['Age']}, \nLocation: {row['Location']}, \nGender: {row['Gender']}\n\n")

    # Save the newly created Word document
    doc.save(output_docx)
    print(f"Word document saved as: {output_docx}")

    # Convert the Word document to PDF
    convert_to_pdf(output_docx, output_pdf)

# Part 3: Word to PDF Conversion
def convert_to_pdf(input_docx, output_pdf):
    wdFormatPDF = 17
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(input_docx)
    doc.SaveAs(output_pdf, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    print(f"PDF document saved as: {output_pdf}")



# Main function to run both parts
if __name__ == "__main__":
    excel_file = 'C:\\Users\\admin\\Desktop\\Book1.xlsx'
    sheet_name = 'Sheet1'
    output_docx = 'C:\\Users\\admin\\Desktop\\updated_document.docx'
    output_pdf = 'C:\\Users\\admin\\Desktop\\output.pdf'

    # Step 1: Read data from Excel
    excel_data = read_excel(excel_file, sheet_name)

    # Step 2: Generate Word report and save as PDF
    generate_word_report(excel_data, output_docx, output_pdf)