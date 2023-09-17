import os
import re
import pandas as pd
import toml

import json
import sys

# Get center names from the input folde xlsx 
def fetch_center_names(input_folder):
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]
    center_names = set()

    for excel_file in excel_files:
        center_name = re.match(r'^([A-Z]+\s\d+)', excel_file)  # Extract the center name from the file name
        if center_name:
            center_names.add(center_name.group(1))

    return center_names

# # Define the sort key function
# def sort_key(reg_no):
#     parts = reg_no.split("-")  # Split the REG. NO. into parts using '-'
#     # print(reg_no)
#     course_number = parts[0]   # Get the course number
#     year_parts = parts[2].split("/")  # Split the year part further using '/'
    
#     year_number = -int(year_parts[1])   # Get the year number ("-" for largest year)
#     student_number = int(year_parts[0]) # Get the student number
    
#     # Return a tuple of values that determine the sorting order
#     return (year_number, student_number, course_number, reg_no)


# Define the sort key function
def sort_key(reg_no):
    parts = re.split(r'[‐-]', reg_no)  # Split the REG. NO. using either regular or special hyphen
    # print(reg_no)
    course_number = parts[0]   # Get the course number
    year_parts = parts[2].split("/")  # Split the year part further using '/'
    
    year_number = -int(year_parts[1])   # Get the year number ("-" for largest year)
    student_number = int(year_parts[0]) # Get the student number
    
    # Return a tuple of values that determine the sorting order
    return (year_number, student_number, course_number, reg_no)


# Function to check if all files are .xlsx
def check_xlsx_files(folder_path):
    for filename in os.listdir(folder_path):
        if not filename.endswith(".xlsx"):
            print(f"Error: {filename} is not a .xlsx file. Please correct and rerun the program.")
            sys.exit(1)

# Function to check if filenames match Unit Codes
def check_filenames_match_units(folder_path, unit_codes):
    mismatched_filenames = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            base_filename = os.path.splitext(filename)[0]
            if base_filename not in unit_codes:
                mismatched_filenames.append(filename)

    if mismatched_filenames:
        print("Error: The following filenames do not match any Unit Code in the JSON file:")
        for filename in mismatched_filenames:
            print(filename)
        print("Please rename these files to match the Unit Codes and rerun the program.")
        sys.exit(1)

# # Function to check if unit codes belong to a single year
# def check_unit_codes_single_year(unit_codes, year_data):
#     year_units_mapping = {}
#     # print(year_data.items())
#     for year, semesters in year_data.items():
#         if isinstance(semesters, list):
#             for unit in semesters:
#                 if isinstance(unit, dict) and "Unit Code" in unit:
#                     unit_code = unit["Unit Code"]
#                     if unit_code in unit_codes:
#                         if unit_code not in year_units_mapping:
#                             year_units_mapping[unit_code] = [year]
#                         else:
#                             year_units_mapping[unit_code].append(year)
#         elif isinstance(semesters, dict):
#             for semester_units in semesters.values():
#                 for unit in semester_units:
#                     if isinstance(unit, dict) and "Unit Code" in unit:
#                         unit_code = unit["Unit Code"]
#                         if unit_code in unit_codes:
#                             if unit_code not in year_units_mapping:
#                                 year_units_mapping[unit_code] = [year]
#                             else:
#                                 year_units_mapping[unit_code].append(year)
    
#     error_unit_codes = [unit_code for unit_code, years in year_units_mapping.items() if len(set(years)) > 1]
    
#     if error_unit_codes:
#         print("Error: Some unit codes belong to multiple years. Please ensure that exam scores are from one year of study and not mixed.")
#         for unit_code in error_unit_codes:
#             print(f"Unit Code: {unit_code} is from years: {', '.join(year_units_mapping[unit_code])}")
#         sys.exit(1)
    
#     return year_units_mapping.popitem()[1][0]




# # Function to check if unit codes belong to a single year
# def check_unit_codes_single_year(unit_codes, year_data):
#     year_units_mapping = {}
    
#     for year, semesters in year_data.items():
#         if isinstance(semesters, list):
#             for unit in semesters:
#                 if isinstance(unit, dict) and "Unit Code" in unit:
#                     unit_code = unit["Unit Code"]
#                     if unit_code in unit_codes:
#                         if unit_code not in year_units_mapping:
#                             year_units_mapping[unit_code] = [year]
#                         else:
#                             year_units_mapping[unit_code].append(year)
#         elif isinstance(semesters, dict):
#             for semester_units in semesters.values():
#                 for unit in semester_units:
#                     if isinstance(unit, dict) and "Unit Code" in unit:
#                         unit_code = unit["Unit Code"]
#                         if unit_code in unit_codes:
#                             if unit_code not in year_units_mapping:
#                                 year_units_mapping[unit_code] = [year]
#                             else:
#                                 year_units_mapping[unit_code].append(year)
    
#     error_unit_codes = [unit_code for unit_code, years in year_units_mapping.items() if len(set(years)) > 1]
    
#     if error_unit_codes:
#         print("Error: Some unit codes belong to multiple years. Please ensure that exam scores are from one year of study and not mixed.")
#         for unit_code in error_unit_codes:
#             print(f"Unit Code: {unit_code} is from years: {', '.join(year_units_mapping[unit_code])}")
#         sys.exit(1)
    
#     # Determine the year from which the units are coming
#     unique_years = set(year for years in year_units_mapping.values() for year in years)
#     if len(unique_years) == 1:
#         return unique_years.pop()
#     else:
#         print("Error: Units are coming from multiple years.")
#         sys.exit(1)



# Function to check if unit codes belong to a single year
def check_unit_codes_single_year(unit_codes, year_data):
    year_units_mapping = {}
    
    for year, semesters in year_data.items():
        if isinstance(semesters, list):
            for unit in semesters:
                if isinstance(unit, dict) and "Unit Code" in unit:
                    unit_code = unit["Unit Code"]
                    if unit_code in unit_codes:
                        if unit_code not in year_units_mapping:
                            year_units_mapping[unit_code] = [year]
                        else:
                            year_units_mapping[unit_code].append(year)
        elif isinstance(semesters, dict):
            for semester_units in semesters.values():
                for unit in semester_units:
                    if isinstance(unit, dict) and "Unit Code" in unit:
                        unit_code = unit["Unit Code"]
                        if unit_code in unit_codes:
                            if unit_code not in year_units_mapping:
                                year_units_mapping[unit_code] = [year]
                            else:
                                year_units_mapping[unit_code].append(year)
    
    error_unit_codes = [unit_code for unit_code, years in year_units_mapping.items() if len(set(years)) > 1]
    
    if error_unit_codes:
        print("Error: Some unit codes belong to multiple years. Please ensure that exam scores are from one year of study and not mixed.")
        for unit_code in error_unit_codes:
            print(f"Unit Code: {unit_code} is from years: {', '.join(year_units_mapping[unit_code])}")
        sys.exit(1)
    
    # Determine the year from which the units are coming
    unique_years = set(year for years in year_units_mapping.values() for year in years)
    if len(unique_years) == 1:
        return unique_years.pop()
    else:
        print("Error: Units are coming from multiple years.")
        for unit_code, years in year_units_mapping.items():
            print(f"Unit Code: {unit_code} is from years: {', '.join(years)}")
        sys.exit(1)