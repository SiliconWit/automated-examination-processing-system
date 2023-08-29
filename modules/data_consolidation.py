import os, toml, re
import pandas as pd
from .file_processing import *
from .utilities import *

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

    # Combine all columns: desired columns, dynamic center columns, additional columns
    new_column_order = desired_columns + center_columns + additional_columns

    global collected_data  # Use the global collected_data list
    # Create an empty DataFrame using the generated column list
    consolidated_df = pd.DataFrame(columns=new_column_order)

    # Loop through each Excel file and consolidate the data
    loop_to_consolidate(excel_files)

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
