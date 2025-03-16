import os
import re
import subprocess
import time
from PyPDF2 import PdfReader
from fpdf import FPDF


def extract_header_file_data(pdf_path):
    """
    Extract Style, Fabric, Version, Color, and Cutting Template from a given PDF file.
    Returns a dictionary with the extracted data.
    """
    reader = PdfReader(pdf_path)
    extracted_data = {}
    first_page = reader.pages[0]
    page_text = first_page.extract_text()

    # Regex patterns
    order_pattern = r"5100[0-9]+"  # Matches order numbers like 5100381797
    # Matches total quantity like 1.509
    total_quantity_pattern = r"Gesamtmenge\s+([\d.,]+)\s+ST"
    # Matches model names like Trousers HestenM204X
    model_pattern = r"(Trousers|Jackets|Suits)\s+([A-Za-z0-9\-/]+)"

    # Extract data using regex
    order_matches = re.findall(order_pattern, page_text)
    total_quantity_matches = re.findall(total_quantity_pattern, page_text)
    model_matches = re.findall(model_pattern, page_text)[0][1]

    total_quantity_matches = [quantity.replace(
        ".", "") for quantity in total_quantity_matches]

    # Store extracted data in a dictionary
    extracted_data = {
        "order_numbers": order_matches[0],
        "total_quantities": total_quantity_matches[0],
        "model_names": model_matches,
    }

    return extracted_data


# def is_printer_busy(printer_name):
    """Check if the printer has pending jobs."""
    cmd = ['wmic', 'printer', 'get', 'Name,JobCount']
    result = subprocess.run(cmd, capture_output=True, text=True)

    for line in result.stdout.splitlines():
        if printer_name in line:
            job_count = int(line.split()[-1])
            return job_count > 0
    return False


def close_adobe():
    """Force close Adobe Acrobat to free the PDF file."""
    subprocess.run(["taskkill", "/F", "/IM", "Acrobat.exe"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def create_header_file(model_name, order_number, total_quantity, output_path):
    """
    Create a PDF header file with bold, large text in landscape orientation.

    Args:
        model_name (str): The name of the model (e.g., "HestenM204X").
        order_number (str): The order number (e.g., "5100381797").
        total_quantity (str): The total quantity (e.g., "1509").
        output_path (str): The path where the PDF will be saved.

    Returns:
        str: The path to the created PDF file.
    """
    # Create a PDF object with landscape orientation
    pdf = FPDF(orientation='L')  # 'L' for landscape
    pdf.add_page()

    pdf.set_font("Arial", style="B", size=60)  # Smaller size for model name
    pdf.cell(0, 30, txt=f"Model: {model_name}", ln=True, align="C")

    pdf.set_font("Arial", style="B", size=72)  # Larger size for order and quantity
    pdf.cell(0, 50, txt=f"Order: {order_number}", ln=True, align="C")
    pdf.cell(0, 50, txt=f"Quantity: {total_quantity}", ln=True, align="C")

    # Save the PDF to the specified output path
    pdf.output(output_path)

    return output_path


def is_printer_busy(printer_name):
    """Check if the printer has pending jobs."""
    cmd = ['wmic', 'printer', 'get', 'Name,JobCount']
    result = subprocess.run(cmd, capture_output=True, text=True)

    for line in result.stdout.splitlines():
        if printer_name in line:
            job_count = int(line.split()[-1])
            return job_count > 0
    return False


def print_file(header_file_name, printer_name):
    """Print a PDF file and ensure it prints before moving to the next."""
    try:
        acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
        command = [acrobat_path, "/t", header_file_name]

        process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        time.sleep(5)  # Allow time for print job to be added to the queue

        while is_printer_busy(printer_name):
            time.sleep(5)  # Reduce sleep time for faster processing

        close_adobe()  # Close Adobe after printing

    except Exception as e:
        print(f"Failed to print {header_file_name}: {e}")


def process_folders(orders_folder, printer_name):
    """Extract data from Folder A, find matching PDFs in Folder B, and print them in order."""
    for file_name in os.listdir(orders_folder):
        if file_name.endswith("PDF"):
            pdf_path_a = os.path.join(orders_folder, file_name)
            extracted_data = extract_header_file_data(pdf_path_a)

            # Extract data for the header file
            model_name = extracted_data["model_names"]
            order_number = extracted_data["order_numbers"]
            total_quantity = extracted_data["total_quantities"]

            # Create the header file
            header_file_path = os.path.join(
                r"C:\Users\Admin\Desktop\scirpts", f"Header_{order_number}.pdf")
            create_header_file(model_name, order_number,
                               total_quantity, header_file_path)

            # Print the header file
            print(f"Printing header file for order {order_number}")
            print_file(header_file_path, printer_name)


# Define folder paths
orders_folder = input("Въведете път към транспорт: ")

# Set your printer name here
printer_name = "YourPrinterName"

# Process the folders
process_folders(orders_folder, printer_name)
