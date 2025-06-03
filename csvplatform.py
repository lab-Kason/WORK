import streamlit as st
import csv
import os
import sys
import tempfile
from PyPDF2 import PdfReader  # For reading PDF files
from docx import Document  # For reading .docx files
import xlrd  # For reading .xls files
from openpyxl import load_workbook  # For reading .xlsx files

# Increase recursion limit
sys.setrecursionlimit(10000)

# File extraction functions
def extract_text_from_pdf(pdf_path):
    try:
        # Attempt extraction with PyPDF2
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        st.error(f"Could not read PDF file {pdf_path}: {e}")
        return ""

def extract_text_from_txt(txt_path):
    try:
        with open(txt_path, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        st.error(f"Could not read TXT file {txt_path}: {e}")
        return ""

def extract_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])
    except Exception as e:
        st.error(f"Could not read DOCX file {docx_path}: {e}")
        return ""

def extract_text_from_xls(xls_path):
    try:
        workbook = xlrd.open_workbook(xls_path)
        data = []
        for sheet in workbook.sheets():
            for row_idx in range(sheet.nrows):
                row = sheet.row_values(row_idx)
                data.append(row)
        return data
    except Exception as e:
        st.error(f"Could not read XLS file {xls_path}: {e}")
        return []

def extract_text_from_xlsx(xlsx_path):
    try:
        workbook = load_workbook(xlsx_path, data_only=True)
        text = ""
        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            for row in worksheet.iter_rows(values_only=True):
                text += " ".join([str(cell) for cell in row if cell is not None]) + "\n"
        return text
    except Exception as e:
        st.error(f"Could not read XLSX file {xlsx_path}: {e}")
        return ""

def extract_text(file_path):
    if file_path.lower().endswith(".pdf"):
        return extract_text_from_pdf(file_path)
    elif file_path.lower().endswith(".txt"):
        return extract_text_from_txt(file_path)
    elif file_path.lower().endswith(".docx"):
        return extract_text_from_docx(file_path)
    elif file_path.lower().endswith(".xls"):
        return extract_text_from_xls(file_path)
    elif file_path.lower().endswith(".xlsx"):
        return extract_text_from_xlsx(file_path)
    else:
        st.error(f"Unsupported file type: {file_path}")
        return ""

# Streamlit app
def main():
    st.title("File Data Extraction and CSV Generator")
    
    # File upload
    uploaded_files = st.file_uploader("Upload files", accept_multiple_files=True)
    
    # Column titles and keywords
    column_titles = st.text_input("Enter column titles (comma-separated)", "Item,Description,Qty,Amount")
    column_titles = [title.strip() for title in column_titles.split(",")]
    
    keywords = {}
    for column in column_titles:
        keywords[column] = st.text_input(f"Enter keyword for column '{column}'", column)
    
    # Extraction behaviors
    extraction_behaviors = {}
    for column in column_titles:
        extraction_behaviors[column] = st.selectbox(
            f"Select extraction behavior for column '{column}'",
            ["right", "left", "below", "above", "keyword"],
            index=2
        )
    
    # Process files and generate CSV
    if st.button("Generate CSV"):
        if uploaded_files:
            rows = []
            for uploaded_file in uploaded_files:
                # Save the uploaded file to a temporary directory
                file_extension = os.path.splitext(uploaded_file.name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                # Extract text and process the file
                text = extract_text(temp_file_path)
                extracted_data = extract_data_from_pdf(text, keywords, extraction_behaviors)
                row = [extracted_data.get(column, "N/A") for column in column_titles]
                rows.append(row)
            
            # Save to CSV
            csv_file_path = os.path.join(os.getcwd(), "output.csv")
            with open(csv_file_path, mode="w", newline="") as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(column_titles)
                writer.writerows(rows)
            
            # Provide download button for the CSV file
            with open(csv_file_path, "rb") as f:
                st.download_button(
                    label="Download CSV",
                    data=f.read(),
                    file_name="output.csv",
                    mime="text/csv"
                )
        else:
            st.error("No files uploaded!")

if __name__ == "__main__":
    main()
