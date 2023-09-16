import os
import pandas as pd

def extract_and_save_all_sheets(input_file_path, output_folder):
    # Read the Excel file
    xls = pd.ExcelFile(input_file_path)
    
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Loop through all sheet names
    for sheet_name in xls.sheet_names:
        # Read the sheet
        df = pd.read_excel(input_file_path, sheet_name=sheet_name)
        
        # Create output file path
        output_file_path = os.path.join(output_folder, f"{sheet_name}.xlsx")
        
        # Save the sheet to a new Excel file
        df.to_excel(output_file_path, index=False)
        print(f"Sheet '{sheet_name}' extracted and saved to '{output_file_path}'.")

if __name__ == "__main__":
    # Specify the file path of the source Excel file and the output folder
    source_excel_file_path = "../../MET_MARCH_2023_YR_4.2_CMS.xlsx"
    output_folder = "../../extracted_sheets"
    
    # Extract and save all sheets
    extract_and_save_all_sheets(source_excel_file_path, output_folder)
    
    print("All sheets extracted and saved.")

