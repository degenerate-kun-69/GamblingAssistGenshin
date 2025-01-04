import pandas as pd
import os

def split_excel_sheets(input_file, output_directory):
    # Load the Excel file
    try:
        excel_file = pd.ExcelFile(input_file)
    except Exception as e:
        print(f"Error loading file: {e}")
        return
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # Iterate through each sheet and save it as a separate Excel file
    for sheet_name in excel_file.sheet_names:
        try:
            df = excel_file.parse(sheet_name)
            output_file = os.path.join(output_directory, f"{sheet_name}.xlsx")
            df.to_excel(output_file, index=False)
            print(f"Sheet '{sheet_name}' saved as '{output_file}'")
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")
    
    print("All sheets have been separated and saved.")

# Usage
input_excel_file = 'wishHistory.xlsx'
output_folder = 'C:/Users/srija/OneDrive/Documents/Programming/projects/Genshin Gamble Assist/GamblingAssistGenshin/preprocesing'
split_excel_sheets(input_excel_file, output_folder)
