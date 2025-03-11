import os
import re
import PyPDF2


def extract_order_number(text):
    # Regex to find a 10-digit number
    order_number_match = re.search(r'\b\d{10}\b', text)
    if order_number_match:
        return order_number_match.group(0)
    return None


def determine_garment_type(text):
    if "Jackets" in text:
        return "Jackets"
    elif "Suits" in text:
        return "Suits"
    elif "Trousers" in text:
        return "Trousers"
    return None


def extract_styles_and_colors(text, style_color_dict):
    styles = ["STA Style D1", "STA Style D", "STA Style E",
              "STA Style Q", "STA Style PT", "STA Style PT1", "STA Style O"]

    for style in styles:
        # Regex to find the style, the 8-digit number, and the numbers (colors) following it
        pattern = re.compile(rf'{style}\s+(\d{{8}})\s+([\d\s]+)')
        matches = pattern.findall(text)
        if matches:
            # Extract the 8-digit number and trim it to the last 4 digits
            full_number = matches[0][0]
            last_four_digits = full_number[-4:]

            # Split the numbers (colors) by any whitespace and store them in a set
            new_colors = set(re.split(r'\s+', matches[0][1].strip()))

            # Create a key for the style and last four digits
            key = f"{style} {last_four_digits}"

            # If the key already exists, update the existing set of colors
            if key in style_color_dict:
                style_color_dict[key].update(new_colors)
            else:
                # Otherwise, create a new set for this style
                style_color_dict[key] = new_colors

            # Debugging: Print the dictionary after updating

    return style_color_dict


def process_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        first_page_text = reader.pages[0].extract_text()

        order_number = extract_order_number(first_page_text)
        garment_type = determine_garment_type(first_page_text)

        if not order_number or not garment_type:
            return None

        result = {
            "order_number": order_number,
            "garment_type": garment_type,
            "styles": {}
        }

        # Initialize the dictionary to accumulate colors across pages
        style_color_dict = {}

        for page in reader.pages:
            page_text = page.extract_text()
            if "Farbzuordnung Komponenten" in page_text:
                # Update the dictionary with colors from the current page
                style_color_dict = extract_styles_and_colors(
                    page_text, style_color_dict)

        # Add the accumulated styles and colors to the result
        result["styles"] = style_color_dict

        return result


def process_folder(folder_path):
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".PDF"):
            file_path = os.path.join(folder_path, filename)
            result = process_pdf(file_path)
            if result:
                results.append(result)
    return results


def print_results(results):
    for result in results:
        print(f"Order Number: {result['order_number']}")
        print(f"Garment Type: {result['garment_type']}")
        for style_with_number, colors in result['styles'].items():
            print(f"{style_with_number}: {', '.join(colors)}")
        print("-" * 40)


# Example usage
folder_path = input("Въведете път до транспорта: ")
results = process_folder(folder_path)
print_results(results)
os.system("pause")
