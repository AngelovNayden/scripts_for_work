import os

def search_models_in_directory(directory, models):
    model_count = {model: 0 for model in models}

    print("\n--- STARTING DIRECTORY SCAN ---\n")
    print(f"Scanning main directory: {directory}\n")

    for root, dirs, files in os.walk(directory):
        # Look for "Сезон" in folder names
        for dir_name in dirs:
            if "сезон" in dir_name.lower():
                season_folder_path = os.path.join(root, dir_name)
                print(f"[INFO] Found season folder: {season_folder_path}")

                for season_root, season_dirs, season_files in os.walk(season_folder_path):
                    # Look for "Транспорт" folders
                    for transport_dir in season_dirs:
                        if "транспорт" in transport_dir.lower():
                            transport_folder_path = os.path.join(season_root, transport_dir)
                            print(f"  ├── [INFO] Found transport folder: {transport_folder_path}")

                            for transport_root, transport_dirs, transport_files in os.walk(transport_folder_path):
                                # Check inside transport folders
                                for subfolder in transport_dirs:
                                    print(f"    ├── [DEBUG] Checking model subfolder: {subfolder}")
                                    for model in models:
                                        if model.lower() in subfolder.lower():
                                            model_count[model] += 1
                                            print(f"      ├── [MATCH] Model '{model}' found in '{subfolder}'")

    print("\n--- SEARCH RESULTS ---\n")
    for model, count in model_count.items():
        if count >= 1 or model == "Henry/Getlin232X":
            print(f"✅ Не е нов модел: {model}")
        else:
            print(f"❌ Нов модел е: {model} !!!")

    print("\n--- SCAN COMPLETE ---\n")

if __name__ == "__main__":
    main_directory = input("Въведете път към транспорта: ").strip()
    models_to_search = []
    print("Въведете моделите (по един на ред). Завършете с празен ред:")
    
    while True:
        model = input().strip()
        if not model:
            break
        models_to_search.append(model)

    search_models_in_directory(main_directory, models_to_search)
    input("\nPress Enter to exit...")
