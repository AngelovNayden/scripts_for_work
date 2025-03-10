import os
import re
import pdfplumber


def extract_fabrics_data(pdf_path):
    """
    For the first page:
      - Extract the garment type (Trousers, Jackets, or Suits).
      - Extract the quantity of garments (with "ST" after the number).
    For pages (from page 2 onward) that contain "Bestellpositionen":
      - Extract the 10-digit order number.
      - Extract the first occurrence of an 8-digit number after "Style".
      - Extract the type of fabric from the line starting with "Techn. Design".
      - Look for the line starting with "Farbe" and capture all three-digit numbers following it.
    Store the data in a nested dictionary structure.
    """
    style_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        # Process the first page for garment type and quantity
        first_page = pdf.pages[0]
        first_page_text = first_page.extract_text()

        # Determine the garment type (Trousers, Jackets, or Suits)
        garment_type = None
        if "Trousers" in first_page_text:
            garment_type = "Trousers"
        elif "Jackets" in first_page_text:
            garment_type = "Jackets"
        elif "Suits" in first_page_text:
            garment_type = "Suits"

        # Extract the quantity of garments (with "ST" after the number)
        quantity_match = re.search(
            r"(\d{1,3}(?:\.\d{3})?)\s*ST", first_page_text)
        if quantity_match:
            # Remove thousand separators (e.g., "1.949" -> "1949")
            quantity = quantity_match.group(1).replace(".", "")
        else:
            quantity = None

        # Process the remaining pages for order number, fabric number, fabric type, and colors
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()
            if not text or "Bestellpositionen" not in text:
                continue

            # Extract the 10-digit order number
            order_number_match = re.search(r"\b\d{10}\b", text)
            if not order_number_match:
                continue  # no order number found on this page
            order_number = order_number_match.group(0)

            # Get the first occurrence of Style with an 8-digit number
            style_match = re.search(r"Qualität\s+(\d{8})", text)
            if not style_match:
                continue  # no style found on this page
            fabric_number = style_match.group(1)

            # Extract the type of fabric from the line starting with "Techn. Design"
            fabric_type_match = re.search(
                r"Techn\. Design\s+([A-Za-z]+)", text)
            fabric_type = fabric_type_match.group(
                1) if fabric_type_match else None

            # Look for a line starting with "Farbe" and capture all three-digit numbers on that line.
            farbe_line_match = re.search(r"Farbe\s+((?:\d{3}\s+)+)", text)
            farbe_matches = re.findall(
                r"\d{3}", farbe_line_match.group(1)) if farbe_line_match else []

            # Update the style_data dictionary
            if order_number not in style_data:
                style_data[order_number] = {
                    "quantity": quantity,
                    "garment_type": garment_type,
                    "type_of_fabric": set(),
                    "fabrics": {}
                }

            # Add the fabric type to the set
            if fabric_type:
                style_data[order_number]["type_of_fabric"].add(fabric_type)

            # Add the fabric number and colors to the nested dictionary
            if fabric_number not in style_data[order_number]["fabrics"]:
                style_data[order_number]["fabrics"][fabric_number] = set(
                    farbe_matches)
            else:
                style_data[order_number]["fabrics"][fabric_number].update(
                    farbe_matches)

    return style_data


if __name__ == "__main__":
    folder1 = input("Въведете път към транспорта: ")

    fabrics_data = {}

    print("Waiting . /")
    print("Waiting .. \\")
    print("Waiting ... /")

    # Process all PDFs in Folder 1
    for filename in os.listdir(folder1):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder1, filename)
            data = extract_fabrics_data(pdf_path)
            # Update the dictionary with non-repeating entries
            for order_number, order_data in data.items():
                if order_number not in fabrics_data:
                    fabrics_data[order_number] = order_data
                else:
                    # Merge fabrics and colors
                    for fabric_number, colors in order_data["fabrics"].items():
                        if fabric_number in fabrics_data[order_number]["fabrics"]:
                            fabrics_data[order_number]["fabrics"][fabric_number].update(
                                colors)
                        else:
                            fabrics_data[order_number]["fabrics"][fabric_number] = colors
                    # Merge fabric types
                    fabrics_data[order_number]["type_of_fabric"].update(
                        order_data["type_of_fabric"])

    # Print the results
    for order_number, order_data in fabrics_data.items():
        print(f"Order Number: {order_number}")
        print(f"Quantity: {order_data['quantity']} ST")
        print(f"Garment Type: {order_data['garment_type']}")
        print(f"Type of Fabric: {', '.join(order_data['type_of_fabric'])}")
        for fabric_number, colors in order_data["fabrics"].items():
            print(
                f"Fabric Number: {fabric_number}, Colors: {', '.join(colors)}")
        print("-" * 40)
    os.system("pause")
