import pandas as pd

def find_header_row(sheet):
    # Iterate through rows to find the header row
    for index, row in sheet.iterrows():
        if "S/N" in row.values and "REG. NO." in row.values and "NAME" in row.values:
            return index
    raise ValueError("Header row not found")

def get_column_names(file_path, sheet_name):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Find the index of the header row
    header_index = find_header_row(df)
    
    # Set the header row as column names
    df.columns = df.iloc[header_index]
    
    # Get column names
    column_names = df.columns.tolist()
    
    return column_names

if __name__ == "__main__":
    # Specify the file path and sheet name
    excel_file_path = "MET_MARCH_2023_YR_4.2_CMS.xlsx"
    sheet_name = "EMT 4201"  # Change this to the name of your sheet
    
    # Get column names
    columns = get_column_names(excel_file_path, sheet_name)
    
    # Print column names
    print("Column Names:")
    for col in columns:
        print(col)

