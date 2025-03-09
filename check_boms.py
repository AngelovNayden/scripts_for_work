import os
import re
import pdfplumber

def extract_first_document_data(pdf_path):
    """
    For pages (from page 2 onward) that contain "Bestellpositionen":
      - Extract the first occurrence of an 8-digit number after "Style"
      - Look for the line starting with "Farbe" and capture all three-digit numbers following it.
    Store unique (style, farbe_list) pairs in a dictionary.
    """
    style_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()
            if not text or "Bestellpositionen" not in text:
                continue

            # Get the first occurrence of Style with an 8-digit number
            style_match = re.search(r"Style\s+(\d{8})", text)
            if not style_match:
                continue  # no style found on this page
            style = style_match.group(1)

            # Look for a line starting with "Farbe" and capture all three-digit numbers on that line.
            farbe_line_match = re.search(r"Farbe\s+((?:\d{3}\s+)+)", text)
            if farbe_line_match:
                # Extract all three-digit numbers from the captured group.
                farbe_matches = re.findall(r"\d{3}", farbe_line_match.group(1))
            else:
                farbe_matches = []

            # Ensure that farbe_matches is unique for this style.
            if style in style_data:
                for farbe in farbe_matches:
                    if farbe not in style_data[style]:
                        style_data[style].append(farbe)
            else:
                style_data[style] = list(set(farbe_matches))
    return style_data

def extract_second_document_data(pdf_path):
    """
    Opens the PDF and returns:
    - A dictionary where the key is the first 8-digit number found in the document,
      and the value is a set of **unique** three-digit numbers on the line containing the header.
    """
    # Distinctive header line to identify the relevant line
    header_pattern = r"No\.\s+Material\s+Group\s+Component\s+Material\s+Description\s+BOM\s+Usage\s+Comment\s+Dec\.Gr\.\s+MRP\s+\+AMF\s+-HUO\s+Cat\.\s+Qty\.\s+Cons\.\s+Cons\.\s+FS\s+FC"
    three_digit_pattern = r"\b\d{3}\b"

    result = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                # Replace any HTML line breaks to ensure proper splitting.
                text = text.replace("<br>", "\n")

                # Look for the first occurrence of an 8-digit number.
                match = re.search(r"\b\d{8}\b", text)
                if match:
                    style_number = match.group(0)
                    # Initialize the set if the style number is not already in the result
                    if style_number not in result:
                        result[style_number] = set()

                # Split the text into lines and look for the header line.
                lines = text.split("\n")
                for line in lines:
                    if re.search(header_pattern, line):
                        # Extract all three-digit numbers from this specific line.
                        three_digit_matches = re.findall(three_digit_pattern, line)
                        if style_number in result:
                            # Add to the set to ensure uniqueness
                            result[style_number].update(three_digit_matches)

    # Convert sets back to lists for consistency with the first document's format
    result = {style: list(farbe) for style, farbe in result.items()}
    return result

def compare_extracted_data(dict1, dict2):
    """
    Compare the data extracted from the first and second documents.
    Print styles and their Farbe that are missing in the second document but exist in the first.
    Also, compare the colors for each style and print if colors are missing in the second document.
    """
    print("Data from first document (Style: Farbe):")
    for style, farbe in dict1.items():
        print(f"Style {style}: Farbe {farbe}")

    print("\nData from second document:")
    for style, farbe in dict2.items():
        print(f"Style {style}: Farbe {farbe}")

    # Find styles that are missing in the second document but exist in the first
    styles_missing_in_second = set(dict1.keys()) - set(dict2.keys())

    if styles_missing_in_second:
        print("\nStyles missing in the second document (but exist in the first):")
        for style in styles_missing_in_second:
            print(f"Style {style}: Farbe {dict1[style]}")
    else:
        print("\nAll styles from the first document are present in the second document.")

    # Compare colors for each style
    print("\nComparing colors for each style:")
    for style, farbe_list1 in dict1.items():
        if style in dict2:
            farbe_list2 = dict2[style]
            missing_colors = set(farbe_list1) - set(farbe_list2)
            if missing_colors:
                print(f"Style {style} is missing colors in the second document: {missing_colors}")
            else:
                print(f"Style {style} has all colors matching in both documents.")
        else:
            print(f"Style {style} is missing in the second document.")

if __name__ == "__main__":
    folder1 = input("Въведете път към транспорта: ")
    folder2 = input("Въведете път към BOM-овете: ")

    first_doc_data = {}
    second_doc_data = {}

    print("Waiting . /")
    print("Waiting .. \\")
    print("Waiting ... /")

    # Process all PDFs in Folder 1
    for filename in os.listdir(folder1):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder1, filename)
            data = extract_first_document_data(pdf_path)
            # Update the dictionary with non-repeating entries
            for key, value in data.items():
                if key not in first_doc_data:
                    first_doc_data[key] = value

    # Process all PDFs in Folder 2
    for filename in os.listdir(folder2):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder2, filename)
            data = extract_second_document_data(pdf_path)
            # Update the dictionary with non-repeating entries
            for key, value in data.items():
                if key in second_doc_data:
                    second_doc_data[key].extend(value)
                    second_doc_data[key] = list(set(second_doc_data[key]))  # Ensure uniqueness
                else:
                    second_doc_data[key] = value

    # Now compare the extracted data
    compare_extracted_data(first_doc_data, second_doc_data)