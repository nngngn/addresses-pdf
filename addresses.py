import os
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import re
import openpyxl

def read_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Read only the first page
            page = pdf.pages[0]
            text = page.extract_text()
            # Skip the first 2 lines
            lines = text.split('\n')[2:]
            text = '\n'.join(lines)
    except:
        print(f"Failed to read PDF: {pdf_path}")
        return {key: "NOT FOUND" for key in ['Name', 'Address Line 1', 'Address Line 2', 'City', 'State', 'ZIP Code', 'File Name']}

    # Initialize dictionary to store extracted data
    info = {
        'Name': '',
        'Address Line 1': '',
        'Address Line 2': '',
        'City': '',
        'State': '',
        'ZIP Code': '',
    }

    lines = text.split('\n')

    # Find Name, Address, City, and State using patterns
    for i, line in enumerate(lines):
        # Pattern for City
        city_pattern = re.compile(r'([A-Za-z\s]+),\s([A-Z]{2})')
        city_match = city_pattern.search(line)

        # Pattern for Address
        address_pattern = re.compile(r'(\d{3,4}\s[^,]+)')
        address_match = address_pattern.search(lines[i-1]) if i > 0 else None

        # If City is found
        if city_match and address_match:
            info['City'] = city_match.group(1).strip()
            info['State'] = city_match.group(2).strip()
            
            # Extract and process Address Line 1
            full_address = address_match.group().strip()
            # Pattern to identify apartment or unit in the address
            apt_unit_pattern = re.compile(r'\b(Apt|Unit|UNIT|APT|Spc)\b\s.*')
            apt_unit_match = apt_unit_pattern.search(full_address)

            # Check for keywords to cut off Address Line 1
            keywords_to_cut_off = ["Health", "Med", "MED", "HEA", "MRN", " 醫", "醫"]
            for keyword in keywords_to_cut_off:
                keyword_position = full_address.find(keyword)
                if keyword_position != -1:
                    full_address = full_address[:keyword_position].strip()
                    break

            if apt_unit_match:
                split_index = apt_unit_match.start()
                info['Address Line 1'] = full_address[:split_index].strip()
                info['Address Line 2'] = full_address[split_index:].strip()
            else:
                info['Address Line 1'] = full_address

            # Extracting Name from the line before Address
            if i > 1:
                name_line = lines[i-2].strip()
                # Modify the regular expression to capture name until specified keywords
                name_line = re.sub(r'\s*(GROUP|PURCHASER|Purchaser|Medical|MRN|MED|購|IDENTIF|Enrole|N.º).*', '', name_line, flags=re.IGNORECASE)
                info['Name'] = name_line

            # Extract ZIP Code
            zip_pattern = re.compile(r'([A-Z]{2})\s(\d{5}(-\d{4})?)')
            zip_match = zip_pattern.search(line)
            if zip_match:
                info['ZIP Code'] = zip_match.group(2).strip()

            break

    return info


def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_selected = filedialog.askdirectory()
    return folder_selected


def get_pdf_files(folder_path):
    pdf_files = []
    subdirs = next(os.walk(folder_path))[1]

    # Define the order for the prefixes
    prefix_order = {"CD": 1, "GLP": 2, "MP3": 3, "NLP": 4, "PDF": 5}

    def sort_key(folder_name):
        # Splitting folder name to get the prefix and date parts
        parts = folder_name.split()
        if len(parts) >= 2 and parts[0] in prefix_order:
            prefix = parts[0]
            match = re.search(r'(\d+)\.(\d+)\.(\d+)', folder_name)
            if match:
                # Using a tuple (Prefix Order, MM, DD, Number) for sorting
                return (prefix_order[prefix], int(match.group(1)), int(match.group(2)), int(match.group(3)))
            else:
                # Handle case where date part does not match
                return (prefix_order[prefix], float('inf'), float('inf'), folder_name)
        else:
            # Handle case where prefix does not match
            return (float('inf'), float('inf'), float('inf'), folder_name)

    subdirs_sorted = sorted(subdirs, key=sort_key)

    # Iterate over the sorted subdirectories
    for subdir in subdirs_sorted:
        subdir_path = os.path.join(folder_path, subdir)
        # Process files in the immediate subdirectory
        for file in os.listdir(subdir_path):
            if file.endswith('.pdf'):
                pdf_files.append(os.path.join(subdir_path, file))
                break  # Only take one PDF per subfolder
    return pdf_files


def create_spreadsheet(data, filename='addresses.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Adding 'File Name' to headers
    headers = ['Name', 'Address Line 1', 'Address Line 2', 'City', 'State', 'ZIP Code', 'File Name']
    ws.append(headers)

    # Writing the data
    for entry in data:
        row = [entry.get(h, "NOT FOUND") for h in headers]
        ws.append(row)

    wb.save(filename)


# main function
def main():
    folder_path = select_folder()
    pdf_files = get_pdf_files(folder_path)

    extracted_data = []
    for pdf_file in pdf_files:
        info = read_pdf(pdf_file)
        info['File Name'] = pdf_file  # Add 'File Name' to the info dictionary
        extracted_data.append(info)


    # Define the spreadsheet filename
    spreadsheet_filename = 'addresses.xlsx'
    create_spreadsheet(extracted_data, spreadsheet_filename)

    # Print the path where the spreadsheet is saved
    spreadsheet_path = os.path.join(os.getcwd(), spreadsheet_filename)
    print(f"Spreadsheet created at {spreadsheet_path}")

if __name__ == "__main__":
    main()
