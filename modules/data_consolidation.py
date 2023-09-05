import os, toml, re
import pandas as pd
from .file_processing import *
from .utilities import *
from collections import Counter

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

config_path = "config.toml"  # Specify the path to your TOML configuration file
# Load the configuration from the TOML file
config = toml.load(config_path)
input_folder_path = config["input_folder"]["path"]

def consolidate_mark_sheet(input_folder_path, output_excel_path, config_path):
    # Get a list of all Excel files in the input folder

    center_names = fetch_center_names(input_folder_path)

    # Get the column order from the configuration
    desired_columns = config["column_order"]["columns"]
    additional_columns = config["additional_columns"]["columns"]

    # Generate the center names dynamically
    center_names = fetch_center_names(config["input_folder"]["path"])
    center_columns = list(center_names)
    # print(center_columns)

    # Combine all columns: desired columns, dynamic center columns, additional columns
    new_column_order = desired_columns + center_columns + additional_columns

    # global collected_data  # Use the global collected_data list
    # global consolidated_df
    # Create an empty DataFrame using the generated column list
    consolidated_df = pd.DataFrame(columns=new_column_order)
    # print(consolidated_df)
    # consolidated_df.to_csv('collected_data.csv', index=False)

    # Define the course_code variable before calling loop_to_consolidate
    # course_code = None

    # Loop through each Excel file and consolidate the data
    course_code = loop_to_consolidate(excel_files, consolidated_df, collected_data)



    # Group collected data by Reg. No.
    grouped_data = {}
    for course, file_course_code, reg_no, name, internal_marks in collected_data:
        # print(collected_data)
        if reg_no in grouped_data:
            grouped_data[reg_no].append((name, file_course_code, internal_marks))
        else:
            grouped_data[reg_no] = [(name, file_course_code, internal_marks)]

    # open('collected_data.txt', 'w').writelines('\n'.join(map(str, collected_data)) + '\n')
    # open('grouped_data.txt', 'w').writelines('\n'.join(map(str, grouped_data.items())) + '\n')
    # print(course_code)
    # Consolidate names for the same Reg. No.
    consolidated_names = {}
    # print(grouped_data.items())



    for student_id, name_code_mark in grouped_data.items():
        # print(grouped_data.items())
        # print(name_code_mark)
        consolidated_names[student_id] = {'Name': '', 'Code': '', 'Marks': []}
        
        name_counts = Counter([name for name, code, mark in name_code_mark if name and not pd.isna(name) and name.strip() != ""])
        if name_counts:
            most_common_name = name_counts.most_common(1)[0][0]
            consolidated_names[student_id]['Name'] = most_common_name

        course_code_name = [code for name, code, mark in name_code_mark if (name or pd.isna(name) or name.strip() == "") and (name == most_common_name or pd.isna(name) or name.strip() == "")]
        if course_code_name:
            consolidated_names[student_id]['Code'] = course_code_name
        
        marks = [mark for name, code, mark in name_code_mark if (name or pd.isna(name) or name.strip() == "") and (name == most_common_name or pd.isna(name) or name.strip() == "")]
        if marks:
            consolidated_names[student_id]['Marks'] += marks

    # print(consolidated_names)

    # course_code = ['EMT 4104', 'EMT 4202', 'EMT 4103', 'EMT 4105', 'EMT 4201', 'EMT 4204', 'EMT 4205', 'EMT 4101', 'EMT 4203', 'EMT 4102']
    # print(len(course_code))
    # Create an empty DataFrame
    consolidated_data_df = pd.DataFrame(columns=['Reg. No.', 'Name'] + course_code)

    # Populate the DataFrame
    for reg_no, data in consolidated_names.items():
        name = data['Name']
        codes = data.get('Code', [])
        marks = data.get('Marks', [])
        
        # Create a dictionary to hold marks for each course code
        code_marks = {code: mark for code, mark in zip(codes, marks)}
        
        # Fill missing exam codes with empty strings and add to the row
        row = [reg_no, name] + [code_marks.get(code, '') for code in course_code]
        
        consolidated_data_df.loc[len(consolidated_data_df)] = row

    # Save the DataFrame to a CSV file
    # consolidated_data_df.to_csv('collected_data.csv', index=False)




    # Sort the consolidated data based on 'Reg. No.'
    consolidated_data_df['Sort Key'] = consolidated_data_df['Reg. No.'].apply(sort_key)
    consolidated_data_df = consolidated_data_df.sort_values(by='Sort Key')

    # Drop the temporary 'Sort Key' column
    consolidated_data_df.drop(columns=['Sort Key'], inplace=True)

    # Add 'Ser. No.' column with a count
    consolidated_data_df['Ser. No.'] = range(1, len(consolidated_data_df) + 1)

    # consolidated_data_df.to_csv('collected_data.csv', index=False)



    # Get units done in semester_4_1 and semester_4_2
    semester_4_1_units = config['semester_order']['semester_4_1']
    semester_4_2_units = config['semester_order']['semester_4_2']

    # Initialize a list to store the rearranged course codes
    rearranged_course_code = []

    # Rearrange course codes based on semester order
    for unit in semester_4_1_units:
        if unit in course_code:
            rearranged_course_code.append(unit)
        else:
            print(f"Unit {unit} was not done in semester_4_1")

    for unit in semester_4_2_units:
        if unit in course_code:
            rearranged_course_code.append(unit)
        else:
            print(f"Unit {unit} was not done in semester_4_2")




    # Combine all columns: desired columns, dynamic center columns, additional columns
    new_column_order = desired_columns + rearranged_course_code + additional_columns


    # Reorder the columns in the DataFrame using the reindex method
    # new_column_order = ['Ser. No.', 'Reg. No.', 'Name'] + course_code + additional_columns
    consolidated_data_df = consolidated_data_df.reindex(columns=new_column_order)


    # consolidated_data_df.to_csv('collected_data.csv', index=False)



    # Create a new Excel file and save the consolidated data
    # consolidated_data_df.to_excel(output_excel_path, index=False)



    # Create a new Excel file and save the consolidated data
    # output_excel_path = 'output.xlsx'
    wb = Workbook()
    ws = wb.active

    # Add data to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(consolidated_data_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

    # Adjust column widths to fit the data
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

    # Save the workbook
    wb.save(output_excel_path)




    print(f"Consolidated mark sheet saved as '{output_excel_path}'.")

    # Print files without "REG. NO." cell
    if files_without_reg_no:
        print("Files without 'REG. NO.' cell:")
        for file in files_without_reg_no:
            print(file)
