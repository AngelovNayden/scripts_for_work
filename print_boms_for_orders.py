import os
import re
import time
import win32api
from PyPDF2 import PdfReader

def extract_bestellpositionen_data(pdf_path):
    """Extract Style, Fabric, Version, and Color from a given PDF file in Folder A."""
    
    reader = PdfReader(pdf_path)
    # Initial structure for extracted_data
    extracted_data = {
        "styles": set(),
        "fabric": set(),
        "version": set(),
        "color": set(),
        "model_name": set(),
        "cutting_template_name": set()
}

    for page in reader.pages:
        page_text = page.extract_text()
        if not page_text:
            continue
        elif "Bestellpositionen" in page_text:
            # Extract and update other fields
            model_name_matches = re.findall(r"Formbezeichnung\s*([A-Za-z0-9/-]+)", page_text)
            extracted_data["model_name"].update(model_name_matches)
            # Extract Cutting Template
            cutting_template_matches = re.findall(r"Schnittvariante\s*([A-Za-z0-9-]+)", page_text)
            extracted_data["cutting_template_name"].update(cutting_template_matches)

            # Extract Style numbers (8 digits)
            style_matches = re.findall(r"Style\s*(\d{8})", page_text)
            extracted_data["styles"].update(style_matches)

            # Extract Fabric numbers (Qualität - 7+ digits)
            fabric_matches = re.findall(r"Qualität\s*(\d{7,})", page_text)
            extracted_data["fabric"].update(fabric_matches)

            # Extract Version numbers (Fertigungsversion - Format 03-0008)
            version_matches = re.findall(r"Fertigungsversion\s*([\d]{2}-\d{4})", page_text)
            # Process version matches to remove leading zeros from the last part
            processed_versions = []
            for match in version_matches:
                split_version = match.split('-')
                last_part = split_version[-1]  # Get the last part
                last_part_int = int(last_part)  # Remove leading zeros
                processed_versions.append(str(last_part_int))  # Add as a string
            
                extracted_data["version"].update(processed_versions)

            # Extract Color names (Farbe)
            color_matches = re.findall(r"Farbe\s*(\d{3})", page_text)
            extracted_data["color"].update(color_matches)

    return extracted_data

def extract_fabric_description(file_name):
    """Extract fabric name or description from the file name (skipping the prefix)."""
    # Adjusted regex to skip the prefix and capture the rest of the filename
    pattern = r"_(\D+)-(\d+)-(\d+)_V\d+_(\d{4}-\d{1,2}-\d{1,2}-\d{4})\.pdf"
    match = re.search(pattern, file_name)
    
    if match:
        fabric_name = match.group(1)  # This will extract the fabric description (e.g., 'P-Hanry-J-WG-233')
        return fabric_name
    return None

def find_matching_pdf(folder_b, extracted_data):
    """Search for a matching PDF in Folder B based on extracted information from Folder A."""
    matched_boms = []
    pdf_files = [f for f in os.listdir(folder_b) if f.endswith(".pdf")]

    for f in pdf_files:
        season = f.split("_")[1]
        date = f.split("_")[6]
        clean_date = date.replace(".pdf", "")
        style = extracted_data["styles"]
        fabric = extracted_data["fabric"]
        version = extracted_data["version"]
        color = extracted_data["color"]
        model_name = extracted_data["model_name"]
        cutting_template =extracted_data["cutting_template_name"]



def print_pdf(pdf_path, file_name):
     """Send a PDF to the printer."""
     print(f"Printing: {file_name}")
     win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
     time.sleep(5)  # Wait for print job to start

def process_folders(folder_a, folder_b):
    """Extract data from Folder A, find matching PDFs in Folder B, and print them in order."""
    for file_name_a in os.listdir(folder_a):
        if file_name_a.endswith(".PDF"):
            print(file_name_a)
            pdf_path_a = os.path.join(folder_a, file_name_a)

        # Extract details from Folder A's PDF
            extracted_data = extract_bestellpositionen_data(pdf_path_a)

        # Print extracted data
            print(f"\nProcessing file: {file_name_a}")
            print(f"Extracted Data: {extracted_data}")

        # Find matching file in Folder B
            matching_pdf_path = find_matching_pdf(folder_b, extracted_data)
        
            if matching_pdf_path:
                print(f"Found matching PDF: {matching_pdf_path}")
            #print_pdf(matching_pdf_path, os.path.basename(matching_pdf_path))
            else:
                print(f"No matching PDF found for {file_name_a}")

            time.sleep(2)  # Small delay before processing the next file

folder_a = r"\\192.168.1.218\HugoBoss\hugo boss\Lieferungen + PM\2024-2\22 Lieferung"
folder_b = r"\\192.168.1.218\HugoBoss\hugo boss\Lieferungen + PM\2024-2\22 Lieferung\BOMs"
process_folders(folder_a, folder_b)
