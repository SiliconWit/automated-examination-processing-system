import os
import re
import toml
import pandas as pd

config_path = "config.toml"  # Specify the path to your TOML configuration file

# Load the configuration from the TOML file
config = toml.load(config_path)

input_folder_path = config["input_folder"]["path"]

# Load the course patterns from the configuration
course_patterns = config["course_patterns"]

excel_files = [file for file in os.listdir(input_folder_path) if file.endswith('.xlsx')]

# List to store collected data
collected_data = []

# Dictionary to track courses found in each Excel file
course_files = {}

# List to track files without "REG. NO." cell
files_without_reg_no = []

def loop_to_consolidate(excel_files):
    # Loop through each Excel file and consolidate the data
    for excel_file in excel_files:
        file_path = os.path.join(input_folder_path, excel_file)
        df = pd.read_excel(file_path, header=None)  # Read without header

        # Search for the cell containing 'REG. NO.'
        get_reg_no_data(df, excel_file)

def get_reg_no_data(df, excel_file):
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
                check_course_pattern(reg_no_value, data, name_value, excel_file)

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


# Check the course pattern for each course
def check_course_pattern(reg_no_value, data, name_value, excel_file):
    matching_course = None
    for course, pattern in course_patterns.items():
        if re.match(pattern, reg_no_value):
            matching_course = course
            break

    if matching_course:
        data.append((reg_no_value, name_value, matching_course))
    else:
        print(f"Anomaly in file '{excel_file}': Reg. No. value '{reg_no_value}' does not match any of the expected course patterns")
