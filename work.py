import streamlit as st
import pdfplumber
import csv
from datetime import datetime
import re
import io

# Define extraction functions (from pdfidentifier(test).py)
def extract_location(text):
    pattern = r"Details:\s*(TC|CMS)\s+(\d{1,3}[A-Z0-9]?)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return f"{match.group(1).upper()} {match.group(2)}"
    return None

def extract_follow_up(text, location):
    if not location:
        return None
    pattern = rf"{re.escape(location)}\s*(.*)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

def extract_work_order_ref(text):
    pattern = r"W\.O\. REF\. 工作單號碼：\s*(WO\d{9}-\d{3})"
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None

def extract_estimated_cost(text):
    pattern = r"ESTIMATED COST 估計費用\s*:\s*HK\$\s*([\d,]+\.\d{2})"
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None

# Streamlit UI
st.title("PDF Identifier and CSV Generator")

# File uploader for PDFs
uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

# Identifier text
identifier = "DATE日期："

# CSV columns
csv_columns = ["File Path", "接CALL時間", " ", "地點", " ", "跟進事項", " ", "W.O. REF. 工作單號碼：", " ", "ESTIMATED COST 估計費用"]

# Process files when the user clicks the button
if st.button("Process PDFs"):
    if uploaded_files:
        rows_to_write = []
        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
                extracted_date = None
                location = None
                follow_up_action = None
                work_order_ref = None
                estimated_cost = None
                for page in pdf.pages:
                    text = page.extract_text()
                    #st.text(f"Extracted text from page:\n{text}")  # Debug: Print extracted text

                    # Extract data
                    if identifier in text:
                        start_index = text.find(identifier) + len(identifier)
                        extracted_date = text[start_index:].strip().split()[0]
                    location = extract_location(text)
                    follow_up_action = extract_follow_up(text, location)
                    work_order_ref = extract_work_order_ref(text)
                    estimated_cost = extract_estimated_cost(text)

                # Format the date and store the row if found
                if extracted_date:
                    try:
                        parsed_date = datetime.strptime(extracted_date, "%d-%b-%Y")
                        formatted_date = parsed_date.strftime("%-d/%-m/%Y")
                        rows_to_write.append([
                            uploaded_file.name,
                            formatted_date,
                            "",
                            location or "",
                            "",
                            follow_up_action or "",
                            "",
                            work_order_ref or "",
                            "",
                            estimated_cost or ""
                        ])
                    except ValueError as e:
                        st.error(f"Error parsing date in file {uploaded_file.name}: {e}")

        # Write rows to a CSV in memory
        if rows_to_write:
            output = io.StringIO()
            writer = csv.writer(output)
            writer.writerow(csv_columns)
            writer.writerows(rows_to_write)
            st.success("PDFs processed successfully!")

            # Provide a download button
            st.download_button(
                label="Download CSV",
                data=output.getvalue(),
                file_name="output.csv",
                mime="text/csv"
            )
        else:
            st.warning("No valid data found in the uploaded PDFs.")
    else:
        st.warning("Please upload at least one PDF file.")
