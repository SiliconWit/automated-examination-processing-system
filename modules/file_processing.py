import os
import re
import pandas as pd
import toml

# Get center names from the input folde xlsx 
def fetch_center_names(input_folder):
    excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]
    center_names = set()

    for excel_file in excel_files:
        center_name = re.match(r'^([A-Z]+\s\d+)', excel_file)  # Extract the center name from the file name
        if center_name:
            center_names.add(center_name.group(1))

    return center_names

# Define the sort key function
def sort_key(reg_no):
    parts = reg_no.split("-")  # Split the REG. NO. into parts using '-'
    
    course_number = parts[0]   # Get the course number
    year_parts = parts[2].split("/")  # Split the year part further using '/'
    
    year_number = -int(year_parts[1])   # Get the year number ("-" for largest year)
    student_number = int(year_parts[0]) # Get the student number
    
    # Return a tuple of values that determine the sorting order
    return (year_number, student_number, course_number, reg_no)