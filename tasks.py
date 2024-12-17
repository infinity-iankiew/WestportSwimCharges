from pypdf import PdfReader
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re

def extract_pdf_invoice_info(folder_path, invoice_filename):
    # vendor_name = ""
    invoice_no = ""
    date = ""
    quantity = ""
    rate = ""
    container_no = ""
    header_found = False

    with open(fr'{folder_path}\{invoice_filename}', 'rb') as input_file:
        pdf_reader = PdfReader(input_file)
        for page_no in range(len(pdf_reader.pages)):
            text = pdf_reader.pages[page_no].extract_text()
            lines = text.split('\n')

            # Extract vendor name from first line
            # if page_no == 0:
            #     vendor_name = lines[0].strip()
            
            for line in lines: 
                if 'No :' in line:
                    invoice_no_match = re.search(r'No\s*:\s*(\d+)', line)
                    if invoice_no_match:
                        invoice_no = invoice_no_match.group(1)

                if 'Date :' in line:
                    date_match = re.search(r'Date\s*:\s*(\d{2}/\d{2}/\d{4})', line)
                    if date_match:
                        date = date_match.group(1)
                
                if 'DESCRIPTION' in line.upper():
                    header_found = True
                    continue    # Skip header line
                
                if header_found:
                    # Check if container_no is empty or blank
                    if not container_no or not container_no.strip():
                        container_info = line.split('-')
                        container_no = container_info[0].strip()
                    
                    parts = line.split()

                    # Check if the first element is a non-negative integer &
                    # Check if there is at least 6 elements in this table row
                    if len(parts)>=6 and parts[0].isdigit():
                        try:
                            # Make sure that the element get for quantity & rate is float
                            quantity = float(parts[-3])
                            rate = float(parts[-2])
                        except:
                            pass
    
    # Convert the extracted date to a datetime object
    if date:
        date_obj = datetime.strptime(date, '%d/%m/%Y')
        # Subtract one month (e.g. 25/9/2024 -> 25/8/2024)
        one_month_earlier = date_obj - relativedelta(months=1)
        # Format the date style
        # date = one_month_earlier.strftime('%d-%m-%Y')   #(user's Sovy date format)
        date = one_month_earlier.strftime('%d/%m/%Y')

    return invoice_no, date, quantity, rate, container_no