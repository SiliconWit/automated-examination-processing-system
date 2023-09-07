import os, toml, re
import pandas as pd
from .file_processing import *
from .utilities import *
from collections import Counter
from .rule_engine import *

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill

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

    # Create an empty DataFrame using the generated column list
    consolidated_df = pd.DataFrame(columns=new_column_order)

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


    # Sort the consolidated data based on 'Reg. No.'
    consolidated_data_df['Sort Key'] = consolidated_data_df['Reg. No.'].apply(sort_key)
    consolidated_data_df = consolidated_data_df.sort_values(by='Sort Key')

    # Drop the temporary 'Sort Key' column
    consolidated_data_df.drop(columns=['Sort Key'], inplace=True)


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

    # Define the columns to consider for checking for missing marks
    columns_to_check = rearranged_course_code

    # Replace empty strings ('' or ' ') with nan in the specified columns
    consolidated_data_df[columns_to_check] = consolidated_data_df[columns_to_check].replace(['', ' '], np.nan)

    # Use dropna to remove rows with missing values in specified columns
    consolidated_data_df = consolidated_data_df.dropna(subset=columns_to_check, how='all')

    # Reset the index after removing rows
    consolidated_data_df = consolidated_data_df.reset_index(drop=True)

    # Add 'Ser. No.' column with a count
    consolidated_data_df['Ser. No.'] = range(1, len(consolidated_data_df) + 1)

    # Combine all columns: desired columns, dynamic center columns, additional columns
    new_column_order = desired_columns + rearranged_course_code + additional_columns


    # Reorder the columns in the DataFrame using the reindex method
    consolidated_data_df = consolidated_data_df.reindex(columns=new_column_order)



    # Calculate TU (Total Units), Total, and Mean for each row
    for index, row in consolidated_data_df.iterrows():
        total_units = 0
        total_marks = 0
        
        # Calculate total units and total marks for the rearranged_course_code columns
        for code in rearranged_course_code:
            if not np.isnan(row[code]):
                total_units += 1
                total_marks += row[code]
        
        # Fill in TU and Total columns
        consolidated_data_df.at[index, 'TU'] = total_units
        consolidated_data_df.at[index, 'Total'] = total_marks
        
        # Calculate and fill Mean column if all units were done
        if total_units == len(rearranged_course_code):
            mean = total_marks / total_units
            consolidated_data_df.at[index, 'Mean'] = "{:.2f}".format(mean)


    calculate_grade(mean)

    # Add Grade column and fill it based on the Mean column
    consolidated_data_df['Grade'] = consolidated_data_df['Mean'].apply(calculate_grade)


    # Define a function to calculate the recommendation and count supplementaries and special cases
    def calculate_recommendation(row):
        supplementaries = [1 for code in rearranged_course_code if
                        (isinstance(row[code], float) and row[code] < 40) ]

        special_cases = [1 for code in rearranged_course_code if
                        isinstance(row[code], (str, float, np.nan)) and (row[code] == '' or pd.isna(row[code]) or (
                                isinstance(row[code], str) and row[code].isspace()))]

        recommendation = []

        if supplementaries:
            recommendation.append(f'SUPP = {sum(supplementaries)} UNIT{"S" if sum(supplementaries) > 1 else ""}')
        if special_cases:
            recommendation.append(f'SPECIAL = {sum(special_cases)} UNIT{"S" if sum(special_cases) > 1 else ""}')

        return ', '.join(recommendation) if recommendation else 'PASS'


    # Add Recommendation column and fill it based on the rearranged_course_code columns
    consolidated_data_df['Recommendation'] = consolidated_data_df.apply(calculate_recommendation, axis=1)






    # Filter the DataFrame to include only students who passed
    passed_students_df = consolidated_data_df[consolidated_data_df['Recommendation'] == 'PASS']

    # Select the 'Reg. No.' and 'Name' columns
    passed_students_list = passed_students_df[['Reg. No.', 'Name']]

    # Save the filtered data to a new .csv file
    passed_students_list.to_csv('passed_students.csv', index=False)







    # Create a new Excel file and save the consolidated data
    wb = Workbook()
    ws = wb.active

    # Define fill colors 
    red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    light_green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")


    # Add data to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(consolidated_data_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

            # Check if the cell contains a pass recommendation or special case and apply the red fill
            if isinstance(value, str) and ('PASS' in value or value == 'PASS'):
                cell.fill = light_green_fill

            # Check for marks below 40 and color them in a more intense red
            if isinstance(value, (int, float)) and value < 40 and ws.cell(row=1, column=c_idx).value in columns_to_check:
                cell.fill = light_blue_fill

            # Check if the cell contains a supplementary recommendation or special case and apply the red fill
            if isinstance(value, str) and ('SUPP' in value or value == 'SUPP'):
                cell.fill = light_blue_fill

            # Check if the cell contains a special recommendation or special case and apply the red fill
            if isinstance(value, str) and ('SPECIAL' in value or value == 'SPECIAL'):
                cell.fill = red_fill

            # Check for empty strings ('' or ' '), spaces, or nan and color them grey
            elif isinstance(value, str) and (value == '' or value.isspace()):
                cell.fill = red_fill
            elif isinstance(value, float) and np.isnan(value):
                cell.fill = red_fill

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
