import os
import pdfplumber
import re


def extract_garment_type(page_text):
    """
    Extract the garment type (Jackets, Suits, or Trousers) from the first page.
    """
    if "Jackets" in page_text:
        return "Jackets"
    elif "Suits" in page_text:
        return "Suits"
    elif "Trousers" in page_text:
        return "Trousers"
    return None


def extract_made_in_count(pdf_path):
    """
    Extract the total count of "US" pieces from the PDF using regex.
    """
    us_count = 0
    with pdfplumber.open(pdf_path) as pdf:
        # Check the first page for garment type
        first_page_text = pdf.pages[0].extract_text()
        garment_type = extract_garment_type(first_page_text)

        # Skip the document if it's Trousers
        if garment_type == "Jackets" or garment_type == "Suits":
            print(f"ПРОПУСКАТ СЕ {garment_type}!!! {pdf_path}")
            return 0

        # Process pages containing "Bestellpositionen"
        for page in pdf.pages:
            text = page.extract_text()
            if text and "Bestellpositionen" in text:
                # Step 1: Find all "Kategorie" lines
                kategorie_matches = re.findall(
                    r"Kategorie\s*([\d\sBGUSCA]+)", text)

                results = []

                for line in kategorie_matches[0].split(" "):
                    # Extract the last part (e.g., "1", "1US")
                    if "1" in line and not "100" in line and not "001" in line and not "101" in line:
                        if "CA" in line or "US" in line or "1" in line:
                            results.append(line)

                # Step 3: Find the "Gesamtsumme - ST -" line and extract counts
                gesamtsumme_match = re.search(
                    r"Gesamtsumme - ST -\s*([\d\.\s]+)", text)
                if gesamtsumme_match:
                    # Extract all numbers (including dots)
                    gesamtsumme_counts = re.findall(
                        r"\d+\.?\d*", gesamtsumme_match.group(1))

                    gesamtsumme_counts = [count.replace(
                        ".", "") for count in gesamtsumme_counts]

                    # Step 4: Compare "Kategorie" entries with counts
                    for i, kategorie in enumerate(results):
                        if i < len(gesamtsumme_counts):
                            # If the Kategorie contains "US", add the count
                            if "US" in kategorie or "CA" in kategorie:
                                # Convert to float to handle dots
                                count = int(gesamtsumme_counts[i])
                                us_count += count
                                print(
                                    f"Made in за поръчка: {count}")
                        else:
                            print(f"No count found for Kategorie: {kategorie}")
    return us_count


def process_folder(folder_path):
    """
    Process all PDF files in the folder and calculate the total "US" count.
    """
    total_us_count = 0
    for filename in os.listdir(folder_path):
        full_path = os.path.join(folder_path, filename)

    # Skip directories
        if os.path.isdir(full_path):
            continue  # Skip to the next file

    # Process only PDF files
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            us_count = extract_made_in_count(pdf_path)
            total_us_count += us_count
            print(f"File: {filename}, Made in бройки: {us_count}")
        # Print separator
            print("-------------------------------------------------")

    print(f"ОБЩ БРОЙ НА MADE IN ЗА ТРАНСПОРТА: {total_us_count}")


if __name__ == "__main__":
    folder_path = input("Въведете път към транспорт: ")
    process_folder(folder_path)
