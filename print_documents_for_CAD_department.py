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
    bestel, oberStof, fabz, summe, verbrauchliste = search_text

    for i, page in enumerate(reader.pages):
        page_text = page.extract_text()
        if i == 0:
            continue
        if bestel in page_text or oberStof in page_text or fabz in page_text or summe in page_text:
            pages.append(i)
        elif verbrauchliste in page_text and "SUM" in page_text:
            pages.append(i)

    return pages


def extract_order_number(reader):
    """Extract the order number from the first page of the PDF."""
    first_page_text = reader.pages[0].extract_text()
    # Matches a 10-digit number
    match = re.search(r"\b\d{10}\b", first_page_text)
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
            if not order_number:
                print(f"Номерът на поръчката не беше намерен в {
                      file_name}. Пропускане...")
                continue

            # Define the temporary file name
            temp_pdf_path = f"{order_number}_temp_print.pdf"

            # Pages 1st to "Bestellpositionen"
            pages_to_print = find_pages_by_text(
                reader, ["Bestellpositionen", "Oberstoff Klassifizierung", "Farbzuordnung Komponenten", "Summe nach Größe + Farbe der Zuschnittkomponenten", "Verbrauchsliste"])

            # ? Adding first page
            pages_to_print.insert(0, 0)

            # Add the last page
            last_page = len(reader.pages) - 1
            if last_page not in pages_to_print:
                pages_to_print.append(last_page)

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

    time.sleep(1)

    # Notify about saved files
    print("Временни файлове за принтирани копия се запават за преглед")

sciprt_folder_path = r"C:\Users\User\Desktop\за принтиране на поръчки"
orders_path = input("Въведете пътя до транспорта: ")

process_pdfs_for_printing(orders_path)


def print_pdf(pdf_path):
    """Process all PDFs in the folder and print specific pages."""
    """Send a PDF to the printer."""
    win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)


for file_name in os.listdir(sciprt_folder_path):
    if file_name.endswith(".pdf"):
        temp_pdf_path = os.path.join(sciprt_folder_path, file_name)
        print(f"Принтира се: {file_name}")

#    print_pdf(temp_pdf_path)
