import os, json 
import re, sys 
import toml
import numpy as np
import pandas as pd

config_path = "config.toml"  # Specify the path to your TOML configuration file

# Load the configuration from the TOML file
config = toml.load(config_path)

input_folder_path = config["input_folder"]["path"]
running_report_path = config["running_report"]["path"]

# Load the course patterns from the configuration
course_patterns = config["course_patterns"]
# course_patterns = config["course_patterns"]["E022"] # Mechatronics 
# print(course_patterns)

excel_files = [file for file in os.listdir(input_folder_path) if file.endswith('.xlsx')]

# List to store collected data
collected_data = []

# Dictionary to track courses found in each Excel file
course_files = {}
course_code = []
course_code_data = []
# file_course_code = []

# List to track files without "REG. NO." cell
files_without_reg_no = []



def loop_to_consolidate(excel_files, consolidated_df, collected_data):
    # Loop through each Excel file and consolidate the data
    for excel_file in excel_files:
        file_path = os.path.join(input_folder_path, excel_file)
        # global file_course_code 
        file_course_code = file_path.split("/")[-1].split(".xlsx")[0]
        # print(file_course_code)
        # df = pd.read_excel(file_path, header=None)  # Read without header

        # Read the Excel sheet into a Pandas DataFrame
        initial_df = pd.read_excel(file_path, header=None)

        # Define a regular expression pattern for the keyword variations
        summary_keyword_pattern = r'summary\s*of\s*results|summary|analysis'

        # Find the index where the keyword pattern appears (case-insensitive)
        split_indices = initial_df[initial_df.apply(lambda row: re.search(summary_keyword_pattern, str(row), re.IGNORECASE), axis=1).notna()].index

        # Check if any indices were found
        if len(split_indices) > 0:
            # Keep only the rows before the first keyword pattern
            split_index = split_indices[0]
            df = initial_df.iloc[:split_index]
        else:
            # Handle the case where the keyword pattern was not found
            print("Summary keyword pattern for {} not found in the DataFrame.".format(file_path))
            print("Please add the keyword 'summary' or delete the file and rerun.")
            sys.exit(1)

        # Search for the cell containing 'REG. NO.'
        get_reg_no_data(df, excel_file, file_course_code)
    # print(course_code)

    # Extract course codes from course_code_data[0] and remove ".xlsx" and "E022"
    course_code = [key.split('.')[0] for key in course_code_data[0].keys()]

    # Print the resulting list
    # print(course_code)
    return course_code



def get_reg_no_data(df, excel_file, file_course_code):
    # Search for the cell containing 'REG. NO.'
    # Improved search resilient to variations and potential errors in the input files

    reg_no_cell = None
    internal_marks_cell = None
    for index, row in df.iterrows():
        lower_row_values = [str(val).lower() for val in row.values]

        # Define a pattern for 'REG. NO.' and 'INTERNAL EXAMINER MARKS /100' variations using regular expressions
        reg_no_pattern = re.compile(r'reg\s*\.?\s*no\s*\.?|reg_no', re.IGNORECASE)


        # Define a regular expression pattern to match all variations
        internal_marks_cell_pattern = re.compile(r'(int\.?|internal)\s*examiner\s*marks\s*/?\s*100', re.IGNORECASE)

        matching_indices = [i for i, val in enumerate(lower_row_values) if reg_no_pattern.search(val)]
        # print(matching_indices)
        internal_marks_indices = [i for i, val in enumerate(lower_row_values) if internal_marks_cell_pattern.search(val)]
        # print(internal_marks_indices)
        if matching_indices: # access when it is fist not empty 
            reg_no_cell = (index, matching_indices[0])
            # print(reg_no_cell)
            internal_marks_cell = (index, internal_marks_indices[0]) if internal_marks_indices else None
            # print(internal_marks_cell)
            # print(excel_file)
            # break # don't break otherwise you'll stop search 
        

    if reg_no_cell is not None and internal_marks_cell is not None:
        reg_no_row = reg_no_cell[0]
        reg_no_col = reg_no_cell[1]
        internal_marks_row = internal_marks_cell[0]
        internal_marks_col = internal_marks_cell[1]

        # Collect data below the 'REG. NO.' cell and corresponding data to the right
        data = []
        for row_idx in range(reg_no_row + 1, len(df)):
            global reg_no_value
            reg_no_value = df.iloc[row_idx, reg_no_col]
            name_value = df.iloc[row_idx, reg_no_col + 1]
            internal_marks = df.iloc[row_idx, internal_marks_col]
            # print(internal_marks)
            if isinstance(reg_no_value, str):
                reg_no_value = reg_no_value.strip()  # Strip whitespace from value


                if isinstance(internal_marks, str):
                    internal_marks = internal_marks.replace("-", "").strip()
                    if internal_marks.isdigit():
                        internal_marks = int(internal_marks)
                    else:
                        internal_marks = np.nan
                elif isinstance(internal_marks, int):
                    pass
                else:
                    # Convert the integer to float (if needed) and handle other non-string cases
                    if not isinstance(internal_marks, (int, float)):
                        internal_marks = np.nan

                # Check the course pattern for each course and add data
                check_course_pattern(reg_no_value, data, name_value, excel_file, internal_marks, file_course_code)
                
        # open('data0.txt', 'w').writelines('\n'.join(map(str, data)) + '\n')

        # Add the collected data to the list, eliminating duplicates
        for course, file_course_code, reg_no, name, internal_marks in data:
            # open('data0.txt', 'w').writelines('\n'.join(map(str, data)) + '\n')
            if (reg_no, name) not in collected_data:
                collected_data.append((course, file_course_code, reg_no, name, internal_marks))
                # open('../collected_data.txt', 'w').writelines('\n'.join(map(str, collected_data)) + '\n')

        # Store courses found in the current file
        course_files[excel_file] = set(course for course, _, _, _, _, in data)
        course_code_data.append(course_files)
        # print(course_files)
        # print(excel_file)
    elif reg_no_cell is None:
        files_without_reg_no.append(excel_file)
    elif internal_marks_cell is None:
        print("Maybe 'INTERNAL EXAMINER MARKS /100' cell is missing in {}.".format(excel_file))

    # Check for mixed courses in each Excel file
    for excel_file, courses in course_files.items():
        if len(courses) > 1:
            print(f"Warning: Excel file '{excel_file}' contains data from multiple courses: {', '.join(courses)}.")

    # print(course_code_data)


# Check the course pattern for each course
def check_course_pattern(reg_no_value, data, name_value, excel_file, internal_marks, file_course_code):
    matching_course = None
    for course, pattern in course_patterns.items():
        # print(course_patterns.items())
        if re.match(pattern, reg_no_value):
            matching_course = course
            break

    if matching_course:
        data.append((matching_course, file_course_code, reg_no_value, name_value, internal_marks))
    else:
        print(f"Anomaly in file '{excel_file}': Reg. No. value '{reg_no_value}' does not match any of the expected course patterns")


def setup_logging():
    # Create or recreate the 'running_reports.txt' file
    with open(running_report_path, 'w') as log_file:
        pass  # This will create an empty file if it doesn't exist or truncate it if it does

def log_print(text):
    # Append the text to 'running_reports.txt'
    with open(running_report_path, 'a') as log_file:
        log_file.write(text + '\n')
    
    # Print the text to the console
    print(text)

# Find unit name given unit code 
def find_unit_name(mechatronics_units_path, unit_code):
    mechatronics_json_data = json.load(open(mechatronics_units_path))
    for year, semesters in mechatronics_json_data.items():
        for semester, units in semesters.items():
            if isinstance(units, list):
                for unit in units:
                    if unit["Unit Code"] == unit_code:
                        return unit["Unit Title"]
            elif isinstance(units, dict):
                for option, option_units in units.items():
                    for unit in option_units:
                        if unit["Unit Code"] == unit_code:
                            return unit["Unit Title"]
    return "Unit not found"

