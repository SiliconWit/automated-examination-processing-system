import os
import pandas as pd

def combine_sheets_to_excel(input_folder, output_file):
    # Get a list of all Excel files in the input folder
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]
    
    # Create a new Excel writer
    excel_writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    # Loop through each Excel file and add it as a new tab
    for excel_file in excel_files:
        file_path = os.path.join(input_folder, excel_file)
        df = pd.read_excel(file_path)
        sheet_name = os.path.splitext(excel_file)[0]  # Use the filename as the sheet name
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
    
    # Save the combined sheets as a new Excel file
    excel_writer.save()
    print(f"All sheets combined and saved as '{output_file}'.")

if __name__ == "__main__":
    # Specify the input folder containing individual sheets and the output Excel file
    input_folder = "extracted_sheets"
    output_excel_file = "combined_sheets.xlsx"
    
    # Combine the sheets and save as a single Excel file with separate tabs
    combine_sheets_to_excel(input_folder, output_excel_file)
    print(f"All sheets combined and saved as '{output_excel_file}'.")


