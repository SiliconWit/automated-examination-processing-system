import os
import re
import pandas as pd
import numpy as np

def consolidate_mark_sheet(input_folder, output_file):
    # Get a list of all Excel files in the input folder
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]

    # Create an empty DataFrame to store consolidated data
    consolidated_df = pd.DataFrame(columns=['Ser. No.', 'Reg. No.', 'Name', 'EMT 4101', 'EMT 4102', 'EMT 4103',
                                           'EMT 4104', 'EMT 4105', 'SMA 2272', 'EMT 4201',
                                           'EMT 4202', 'EMT 4203', 'EMT 4204', 'EMT 4205',
                                           'SMA 3261', 'TU', 'Total', 'Mean', 'Recommendation'])

    # List to track files without "REG. NO." cell
    files_without_reg_no = []

    # List to store collected data
    collected_data = []

    # Loop through each Excel file and consolidate the data
    for excel_file in excel_files:
        file_path = os.path.join(input_folder, excel_file)
        df = pd.read_excel(file_path, header=None)  # Read without header

        # Search for the cell containing 'REG. NO.'
        # Improved search resilient to variations and potential errors in the input files
        reg_no_cell = None
        for index, row in df.iterrows():
            lower_row_values = [str(val).lower() for val in row.values]

            # Define a pattern for 'REG. NO.' variations using regular expressions
            reg_no_pattern = re.compile(r'reg\s*\.?\s*no\s*\.?|reg_no', re.IGNORECASE)

            matching_indices = [i for i, val in enumerate(lower_row_values) if reg_no_pattern.search(val)]
            if matching_indices:
                reg_no_cell = (index, matching_indices[0])
                break

        if reg_no_cell is not None:
            reg_no_row = reg_no_cell[0]
            reg_no_col = reg_no_cell[1]

            # Collect data below the 'REG. NO.' cell and corresponding data to the right
            data = []
            for row_idx in range(reg_no_row + 1, len(df)):
                reg_no_value = df.iloc[row_idx, reg_no_col]
                name_value = df.iloc[row_idx, reg_no_col + 1]
                if isinstance(reg_no_value, str):
                    reg_no_value = reg_no_value.strip()  # Strip whitespace from value
                    if re.match(r'^E022-01-\d+/\d{4}$', reg_no_value):
                        data.append((reg_no_value, name_value))  # Collect corresponding data
                    else:
                        print(f"Anomaly in file '{excel_file}': Reg. No. value '{reg_no_value}' does not match the expected format")

            # Add the collected data to the list, eliminating duplicates
            for reg_no, name in data:
                if (reg_no, name) not in collected_data:
                    collected_data.append((reg_no, name))
        else:
            files_without_reg_no.append(excel_file)

    # Remove empty values from the collected data
    collected_data = [(reg_no, name) for reg_no, name in collected_data if reg_no and isinstance(name, str)]

    # Group collected data by Reg. No.
    grouped_data = {}
    for reg_no, name in collected_data:
        if reg_no in grouped_data:
            grouped_data[reg_no].add(name)
        else:
            grouped_data[reg_no] = {name}

    # Consolidate names for the same Reg. No.
    consolidated_names = {}
    for reg_no, name_set in grouped_data.items():
        unique_names = sorted(name_set, key=lambda name: name.lower())
        consolidated_names[reg_no] = unique_names[-1] if unique_names else None

    # Apply consolidated names to the collected data
    consolidated_data = [(reg_no, consolidated_names[reg_no]) for reg_no in sorted(grouped_data)]

    # Create a DataFrame from the consolidated data
    consolidated_data_df = pd.DataFrame(consolidated_data, columns=['Reg. No.', 'Name'])

    # Sort the consolidated data based on 'Reg. No.'
    consolidated_data_df['Sort Key'] = consolidated_data_df['Reg. No.'].apply(sort_key)
    consolidated_data_df = consolidated_data_df.sort_values(by='Sort Key')

    # Drop the temporary 'Sort Key' column
    consolidated_data_df.drop(columns=['Sort Key'], inplace=True)

    # Add 'Ser. No.' column with a count
    consolidated_data_df['Ser. No.'] = range(1, len(consolidated_data_df) + 1)

    # Create a new column order that includes the columns in consolidated_data_df
    new_column_order = ['Ser. No.', 'Reg. No.', 'Name'] + list(consolidated_data_df.columns[3:])

    # Reorder the columns in the DataFrame
    consolidated_data_df = consolidated_data_df[new_column_order]

    # Create a new Excel file and save the consolidated data
    consolidated_data_df.to_excel(output_file, index=False)
    print(f"Consolidated mark sheet saved as '{output_file}'.")

    # Print files without "REG. NO." cell
    if files_without_reg_no:
        print("Files without 'REG. NO.' cell:")
        for file in files_without_reg_no:
            print(file)

# Define the sort key function
def sort_key(reg_no):
    parts = reg_no.split("-")  # Split the REG. NO. into parts using '-'
    
    course_number = parts[0]   # Get the course number
    year_parts = parts[2].split("/")  # Split the year part further using '/'
    
    year_number = -int(year_parts[1])   # Get the year number ("-" for largest year)
    student_number = int(year_parts[0]) # Get the student number
    
    # Return a tuple of values that determine the sorting order
    return (year_number, student_number, course_number, reg_no)

if __name__ == "__main__":
    # Specify the input folder containing individual sheets and the output Excel file
    input_folder = "../../individual_sheets"
    output_excel_file = "../../EMT4_2_2023.xlsx"

    # Consolidate the mark sheet and save as a new Excel file
    consolidate_mark_sheet(input_folder, output_excel_file)

    print(f"Mark sheet consolidated and saved as '{output_excel_file}'.")