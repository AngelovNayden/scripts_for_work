import pandas as pd
import os

# Function to create the dictionary from the Excel file


def create_order_delivery_dict(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    df.iloc[:, 4] = df.iloc[:, 4].fillna(method='ffill')

    # Initialize an empty dictionary
    order_delivery_dict = {}

    # Iterate through the rows of column E (assuming E is the 5th column, index 4)
    for index, row in df.iterrows():
        order_number = row[1]  # Column E is index 4 (0-based index)
        delivery_number = row[0]

        sample_order_number = str(row[5]).replace('.0', "")
        sample_delivery_number = row[4]

        # Add to dictionary
        order_delivery_dict[f"{order_number}"] = delivery_number
        order_delivery_dict[sample_order_number] = sample_delivery_number
    return order_delivery_dict


def get_multiline_input():
    print("Въведете номера на поръчки:")
    lines = []
    while True:
        line = input()
        if line.strip() == "":  # Stop when the user enters an empty line
            break
        lines.append(line)
    return "\n".join(lines)


# Function to process the input string and print delivery numbers
def print_delivery_numbers(input_string, order_delivery_dict):
    # Split the input string by new lines to get individual order numbers
    order_numbers = input_string.strip().split('\n')

    # Print the delivery numbers for each order number
    for order in order_numbers:
        order = order.strip()
        if order in order_delivery_dict:
            print(
                f"Транспорт {order_delivery_dict[order]}")
        else:
            print("-")


# Example usage
if __name__ == "__main__":
    # Path to the Excel file
    # Replace with your actual file path
    path_to_excel_file = input("Въведете път до excel-ски файл: ")

    # Create the dictionary
    order_delivery_dict = create_order_delivery_dict(path_to_excel_file)

    # Example input string with multiple order numbers
    input_string = get_multiline_input()

    # Print the delivery numbers

    print_delivery_numbers(input_string, order_delivery_dict)

    os.system("pause")
