import os
import pandas as pd
import numpy as np
import re

def consolidate_mark_sheet(input_folder, output_file):
    # Get a list of all Excel files in the input folder
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]
    
    # Create an empty DataFrame to store consolidated data
    consolidated_df = pd.DataFrame(columns=['Ser. No.', 'Reg. No.', 'EMT 4101', 'EMT 4102', 'EMT 4103',
                                           'EMT 4104', 'EMT 4105', 'SMA 2272', 'EMT 4201',
                                           'EMT 4202', 'EMT 4203', 'EMT 4204', 'EMT 4205',
                                           'SMA 3261', 'TU', 'Total', 'Mean', 'Recommendation'])
    
    # List to track files without "REG. NO." cell
    files_without_reg_no = []
    
    # Set to store unique 'REG. NO.' values
    unique_reg_no_values = set()
    
    # Loop through each Excel file and consolidate the data
    for excel_file in excel_files:
        file_path = os.path.join(input_folder, excel_file)
        df = pd.read_excel(file_path, header=None)  # Read without header
        
        # Search for the cell containing 'REG. NO.'
        reg_no_cell = None
        for index, row in df.iterrows():
            if 'REG. NO.' in row.values:
                reg_no_cell = (index, row.values.tolist().index('REG. NO.'))
                break
        
        if reg_no_cell is not None:
            reg_no_row = reg_no_cell[0]
            reg_no_col = reg_no_cell[1]
            
            # Collect data below the 'REG. NO.' cell
            data = []
            for row_idx in range(reg_no_row + 1, len(df)):
                value = df.iloc[row_idx, reg_no_col]
                if isinstance(value, str):
                    value = value.strip()  # Strip whitespace from value
                    if re.match(r'^E022-01-\d+/\d{4}$', value):
                        data.append(value)
                    else:
                        print(f"Anomaly in file '{excel_file}': Reg. No. value '{value}' does not match the expected format")
                else:
                    data.append(value)
            
            # Add the collected data to the set
            unique_reg_no_values.update(data)
        else:
            files_without_reg_no.append(excel_file)
    
    # Remove empty values and sort unique 'REG. NO.' values
    unique_reg_no_values.discard(np.nan)
    
    # Sort 'REG. NO.' values based on specified criteria
    def sort_key(reg_no):
        parts = reg_no.split("-")  # Split the REG. NO. into parts using '-'
        
        course_number = parts[0]   # Get the course number
        year_parts = parts[2].split("/")  # Split the year part further using '/'
        
        year_number = -int(year_parts[1])   # Get the year number ("-" for largest year)
        student_number = int(year_parts[0]) # Get the student number
        
        # Return a tuple of values that determine the sorting order
        return (year_number, student_number, course_number, reg_no)
    
    sorted_reg_no_values = sorted(unique_reg_no_values, key=sort_key)
    
    # Create a dictionary with sorted 'Reg. No.' values
    reg_no_dict = {'Reg. No.': sorted_reg_no_values}
    
    # Create a DataFrame from the dictionary
    reg_no_df = pd.DataFrame(reg_no_dict)
    
    # Merge the 'Reg. No.' DataFrame with the consolidated DataFrame
    consolidated_df = pd.merge(consolidated_df, reg_no_df, on='Reg. No.', how='outer')
    
    # Remove rows with NaN in 'Reg. No.' column
    consolidated_df.dropna(subset=['Reg. No.'], inplace=True)
    
    # Add a "Ser. No." column with a count
    consolidated_df['Ser. No.'] = range(1, len(consolidated_df) + 1)
    
    # Rearrange the columns
    column_order = ['Ser. No.', 'Reg. No.', 'EMT 4101', 'EMT 4102', 'EMT 4103',
                    'EMT 4104', 'EMT 4105', 'SMA 2272', 'EMT 4201',
                    'EMT 4202', 'EMT 4203', 'EMT 4204', 'EMT 4205',
                    'SMA 3261', 'TU', 'Total', 'Mean', 'Recommendation']
    consolidated_df = consolidated_df[column_order]
    
    # Create a new Excel file and save the consolidated data
    consolidated_df.to_excel(output_file, index=False)
    print(f"Consolidated mark sheet saved as '{output_file}'.")
    
    # Print files without "REG. NO." cell
    if files_without_reg_no:
        print("Files without 'REG. NO.' cell:")
        for file in files_without_reg_no:
            print(file)

if __name__ == "__main__":
    # Specify the input folder containing individual sheets and the output Excel file
    input_folder = "individual_sheets"
    output_excel_file = "EMT4_2_2023.xlsx"
    
    # Consolidate the mark sheet and save as a new Excel file
    consolidate_mark_sheet(input_folder, output_excel_file)
    
    print(f"Mark sheet consolidated and saved as '{output_excel_file}'.")

