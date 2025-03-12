import os
import re
import win32api
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import time


			
def create_dot_page(writer):
    """Add a single page with a small dot to the PDF writer."""
    from reportlab.lib.pagesizes import A4
    from io import BytesIO

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.drawString(10, 10, ".")  # Place a small dot at the bottom-left corner
    c.showPage()
    c.save()

    buffer.seek(0)
    dot_reader = PdfReader(buffer)
    writer.add_page(dot_reader.pages[0])


def find_pages_by_text(reader, search_text):
    """
    Find a range of pages starting from `search_text` and optionally stopping at `stop_text`.
    Excludes pages with `exclude_text` in their content.
    """
    pages = []
    bestel, oberStof, fabz = search_text

    for i, page in enumerate(reader.pages):
        page_text = page.extract_text()
        if i == 0:
            continue
        if bestel in page_text or oberStof in page_text:
            pages.append(i)
        elif fabz in page_text:
            return pages

    return pages


def extract_order_number(reader):
    """Extract the order number from the first page of the PDF."""
    first_page_text = reader.pages[0].extract_text()
    # Matches a 10-digit number
    match = re.search(r"\b\d{10}\b", first_page_text)
    if match:
        return match.group(0)
    return None

def extract_garment(reader):
    """Extract the order number from the first page of the PDF."""
    first_page_text = reader.pages[0].extract_text()
    # Matches the garment
    match = re.search(r"\b(Jackets|Trousers|Suits)\b", first_page_text, re.IGNORECASE)
    if match:
        return match.group(0)
    return None

def process_pdfs_for_printing(folder_path):

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".PDF"):
            pdf_path = os.path.join(folder_path, file_name)
            reader = PdfReader(pdf_path)

            # Extract the order number from the first page
            order_number = extract_order_number(reader)
            garmnent = extract_garment(reader).lower()
            if not order_number:
                print(f"Номерът на поръчката не беше намерен в {
                      file_name}. Пропускане...")
                continue

            # Define the temporary file name
            temp_pdf_path = f"{order_number}_temp_print.pdf"

            # Pages 1st to "Bestellpositionen"
            pages_to_print = find_pages_by_text(
                reader, ["Bestellpositionen", "Oberstoff Klassifizierung", "Farbzuordnung Komponenten"])

            # ? Adding first page
            pages_to_print.insert(0, 0)

            # Remove duplicates and sort the page numbers
            pages_to_print = sorted(set(pages_to_print))

            if not pages_to_print:
                print(f"Няма страници за принтиране за файл {
                      file_name}. Пропускане...")
                continue

            # Add selected pages to the writer
            writer = PdfWriter()
            for page_num in pages_to_print:
                writer.add_page(reader.pages[page_num])

            # Add a blank page with a small dot as the last page
            create_dot_page(writer)

            # Write the temporary PDF
            with open(temp_pdf_path, "wb") as temp_pdf:
                writer.write(temp_pdf)

            print(f"Направено е копие {
                  file_name} и е запазено като {temp_pdf_path}.")
            
            if(garmnent == "trousers"):
                print("Поръчката е от панталони")
            elif garmnent == "jackets":
                print("Поръчката е от сака")
            elif garmnent == "suits":
                print("Поръчката е от костюми")

    time.sleep(1)

    # Notify about saved files
    print("Временни файлове за принтирани копия се запават за преглед")


# Set the folder path containing your PDFs
copies = 0
key_words = ["jackets", "trousers", "suits", "fototeil-bestellung", "muster-bestellung"]

orders_path = input("Въведете път до транспорта: ")
sciprt_folder_path = r"C:\Users\Admin\Desktop\scirpts"


process_pdfs_for_printing(orders_path)

def print_pdf(pdf_path, file_name, copies):
    """Process all PDFs in the folder and print specific pages."""
    """Send a PDF to the printer."""
    for i in range(copies):
        if i == 0:
            print(f"Принтира се 1во копие")
        elif i == 1:
            print(f"Принтира се 2во копие")
        elif i == 2:
            print(f"Принтира се 3то копие")
        win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
    print(f"Принтираха се {copies} копия от {file_name}")


for file_name in os.listdir(sciprt_folder_path):
    if file_name.endswith(".pdf"):
        temp_pdf_path = os.path.join(sciprt_folder_path, file_name)
        print(f"Принтира се: {file_name}")

        reader = PdfReader(temp_pdf_path)
        # Extract text from the first page
        first_page_text = reader.pages[0].extract_text().lower()

        if (key_words[3] in first_page_text or key_words[4] in first_page_text) and (key_words[0] in first_page_text or key_words[1] in first_page_text):
            copies = 1
        elif key_words[2] in first_page_text and (key_words[3] in first_page_text or key_words[4] in first_page_text):
            copies = 2
        elif key_words[0] in first_page_text or key_words[1] in first_page_text:
            copies = 2
        elif key_words[2] in first_page_text:
            copies = 3

        print_pdf(temp_pdf_path, file_name, copies)
