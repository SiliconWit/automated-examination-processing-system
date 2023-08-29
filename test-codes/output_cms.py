import os
import re
import pandas as pd
import numpy as np
import toml

def fetch_center_names(input_folder):
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]
    center_names = set()

    for excel_file in excel_files:
        center_name = re.match(r'^([A-Z]+\s\d+)', excel_file)  # Extract the center name from the file name
        if center_name:
            center_names.add(center_name.group(1))

    return center_names

def consolidate_mark_sheet(input_folder_path, output_excel_path, config_path):
    # Get a list of all Excel files in the input folder

    excel_files = [file for file in os.listdir(input_folder_path) if file.endswith('.xlsx')]

    center_names = fetch_center_names(input_folder_path)

    # Load the configuration from the TOML file
    config = toml.load(config_path)

    # Get the column order from the configuration
    desired_columns = config["column_order"]["columns"]
    additional_columns = config["additional_columns"]["columns"]

    # Generate the center names dynamically
    center_names = fetch_center_names(config["input_folder"]["path"])
    center_columns = list(center_names)

    # Combine all columns: desired columns, dynamic center columns, additional columns
    new_column_order = desired_columns + center_columns + additional_columns

    # Create an empty DataFrame using the generated column list
    consolidated_df = pd.DataFrame(columns=new_column_order)

    # List to track files without "REG. NO." cell
    files_without_reg_no = []

    # Load the course patterns from the configuration
    course_patterns = config["course_patterns"]

    # List to store collected data
    collected_data = []

    # Dictionary to track courses found in each Excel file
    course_files = {}

    # Loop through each Excel file and consolidate the data
    for excel_file in excel_files:
        file_path = os.path.join(input_folder_path, excel_file)
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

                    # Check the course pattern for each course
                    matching_course = None
                    for course, pattern in course_patterns.items():
                        if re.match(pattern, reg_no_value):
                            matching_course = course
                            break

                    if matching_course:
                        data.append((reg_no_value, name_value, matching_course))
                    else:
                        print(f"Anomaly in file '{excel_file}': Reg. No. value '{reg_no_value}' does not match any of the expected course patterns")


            # Add the collected data to the list, eliminating duplicates
            for reg_no, name, course in data:
                if (reg_no, name) not in collected_data:
                    collected_data.append((reg_no, name, course))
            # Store courses found in the current file
            course_files[excel_file] = set(course for _, _, course in data)
        else:
            files_without_reg_no.append(excel_file)
        # Check for mixed courses in each Excel file
        for excel_file, courses in course_files.items():
            if len(courses) > 1:
                print(f"Warning: Excel file '{excel_file}' contains data from multiple courses: {', '.join(courses)}.")


    # Remove empty values from the collected data
    collected_data = [(reg_no, name, course) for reg_no, name, course in collected_data if reg_no and isinstance(name, str)]

    # Group collected data by Reg. No.
    grouped_data = {}
    for reg_no, name, course in collected_data:
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

    # Reorder the columns in the DataFrame using the reindex method
    consolidated_data_df = consolidated_data_df.reindex(columns=new_column_order)

    # Create a new Excel file and save the consolidated data
    consolidated_data_df.to_excel(output_excel_path, index=False)
    print(f"Consolidated mark sheet saved as '{output_excel_path}'.")

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
    config_path = "config.toml"  # Specify the path to your TOML configuration file

    # Load paths from configuration
    config = toml.load(config_path)
    input_folder_path = config["input_folder"]["path"]
    output_excel_path = config["output_excel"]["path"]

    consolidate_mark_sheet(input_folder_path, output_excel_path, config_path)

    print(f"Mark sheet consolidated and saved as '{output_excel_path}'.")