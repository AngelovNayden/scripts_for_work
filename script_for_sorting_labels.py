import os

# Network path with raw string notation
folder_path = r"\\192.168.1.218\HugoBoss\hugo boss\PFLEGEETIKETTEN\29 сезон\22 транспорт"

# Get list of files in the folder
files = os.listdir(folder_path)

# Function to extract the order number and the number of labels from the file name
def extract_order_and_labels(file_name):
    parts = file_name.split("_")
    if len(parts) >= 5:
        order_number = parts[3]  # The 4th element (index 3) is the order number
        labels_count = parts[-1]  # The last element (index -1) is the number of labels
        return order_number, labels_count
    else:
        return None, None  # If the filename doesn't match the expected format

# Rename files based on the new format (order_number_number_of_labels)
for file in files:
    old_path = os.path.join(folder_path, file)
    
    # Extract order number and number of labels from the file name
    order_number, labels_count = extract_order_and_labels(file)
    
    if order_number and labels_count:
        # Get the file extension
        file_extension = os.path.splitext(file)[1]
        
        # Create the new name using the order number and labels count, but without the extension
        new_name = f"{order_number}_{labels_count}"
        
        # Generate the full new file path
        new_path = os.path.join(folder_path, new_name)
        
        # Rename the file in the directory
        os.rename(old_path, new_path)
        print(f"Renamed '{file}' to '{new_name}{file_extension}'")
    else:
        print(f"Skipping '{file}' (filename doesn't match expected format)")
