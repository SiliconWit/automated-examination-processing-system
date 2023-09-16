# import re

# # Define a regular expression pattern to match all variations
# internal_marks_cell_pattern = re.compile(r'(int\.?|internal)\s*examiner\s*marks\s*/?\s*100', re.IGNORECASE)

# # Test the pattern with sample variations
# test_strings = [
#     "internal examiner marks /100",
#     "INTERNAL EXAMINER MARKS  /100",
#     "INT.  EXAMINER MARKS  /100",
# ]

# for test_string in test_strings:
#     if internal_marks_cell_pattern.search(test_string):
#         print(f"Matched: {test_string}")
#     else:
#         print(f"Not Matched: {test_string}")



# import pandas as pd

# # Sample DataFrame
# data = {'Student Name': ['Alice', 'Bob', 'Charlie'],
#         'Recommendation': ['SUPP = 1 UNIT', 'PASS', 'SUPP = 3 UNITS']}
# consolidated_data_df = pd.DataFrame(data)

# # Custom function to check if the Recommendation contains 'SUPP' and units information
# def has_supp_units(recommendation):
#     return 'SUPP' in recommendation and any(word.isdigit() for word in recommendation.split())

# # Filter the DataFrame using the custom function
# supp_students_df = consolidated_data_df[consolidated_data_df['Recommendation'].apply(has_supp_units)]

# # Print the filtered DataFrame
# print(supp_students_df)




# import pandas as pd

# # Sample DataFrame
# data = {'Student Name': ['Alice', 'Bob', 'Charlie', 'Kim'],
#         'Recommendation': ['SUPP = 1 UNIT', 'PASS', 'SUPP = 3 UNITS', 'SPECIAL = 2 UNITS']}
# consolidated_data_df = pd.DataFrame(data)

# # Custom function to check if the Recommendation contains 'SUPP' and does not contain 'SPECIAL' and units information
# def has_supp_units(recommendation):
#     return 'SUPP' in recommendation and 'SPECIAL' not in recommendation and any(word.isdigit() for word in recommendation.split())

# # Filter the DataFrame using the custom function
# supp_students_df = consolidated_data_df[consolidated_data_df['Recommendation'].apply(has_supp_units)]

# # Print the filtered DataFrame
# print(supp_students_df)




import pandas as pd
import numpy as np

grouped_data = [
    ('E022-01-0804/2015', [('Jim Murimi', 'EMT 4104', 72), ('Jim Murimi', 'EMT 4202', 49), ('Jim Murimi', 'EMT 4103', 52), ('Jim Murimi', 'EMT 4105', 50), ('Jim Rochester Murimi', 'EMT 4201', 63), ('Jim Murimi', 'EMT 4204', 58), ('Jim Murimi', 'EMT 4205', 65), ('Jim Murimi', 'EMT 4101', 59), ('Rochester Murimi', 'SMA 2272', 60), ('Jim Rochester Murimi', 'EMT 4203', 67), ('Jim Murimi', 'SMA 3261', 42), ('Jim Murimi', 'EMT 4102', 40)]),
    ('E022-01-1077/2018', [('nan', 'EMT 4104', 76), ("Francis Ng'ang'a", 'EMT 4202', 64), ("Francis Ng'ang'a", 'EMT 4103', 66), ("Francis Ng'ang'a", 'EMT 4105', 50), ("Francis Ng'ang'a", 'EMT 4201', 68), ("Francis Ng'ang'a", 'EMT 4204', 56), ("Francis Ng'ang'a", 'EMT 4205', 47), ('Francis Ngugi Nganga', 'EMT 4101', 27), ('Francis Ngugi', 'SMA 2272', 44), ('Francis Ngugi Nganga', 'EMT 4203', 49), ("Francis Ng'ang'a", 'SMA 3261', 46), ('Francis Ngugi Nganga', 'EMT 4102', 41)]),
    ('E022-01-0805/2019', [('BLAIR CARSON KIPROTICH ', 'EMT 4104', 66), ('Kiprotich Blair Carson', 'EMT 4202', 46), ('Kiprotich Blair Carson', 'EMT 4103', 59), ('Kiprotich Blair Carson', 'EMT 4105', 52), ('Brian Carson KIPROTICH', 'EMT 4201', 70), ('Kiprotich Blair Carson', 'EMT 4204', 66), ('Kiprotich Blair Carson', 'EMT 4205', 55), ('Blair Carson Kiprotich ', 'EMT 4101', 35), ('Blair Carson', 'SMA 2272', 40), ('Blair Carson', 'EMT 4203', 65), ('Blair Carson Kiprotich', 'SMA 3261', 66), ('Blair Carson Kiprotich', 'EMT 4102', 37)]),
]

# Step 1: Standardize names
name_dict = {}
for reg_no, name_score_list in grouped_data:
    for name, _, _ in name_score_list:
        if isinstance(name, str) and len(name.strip()) > 0:
            name = name.title()
            name = ' '.join(word.capitalize() if i != len(name.split()) - 1 else word.upper() for i, word in enumerate(name.split()))
            name_dict[reg_no] = name

# Step 2: Create a list of dictionaries
student_data = []
for reg_no, name_score_list in grouped_data:
    for _, unit_code, score in name_score_list:
        standardized_name = name_dict[reg_no]
        if standardized_name.lower() == 'nan':
            standardized_name = None
        student_data.append({'Registration Number': reg_no, 'Student Name': standardized_name, unit_code: score})

# Step 3: Create a DataFrame
df = pd.DataFrame(student_data)

# Step 4: Pivot the DataFrame to have one row per student
df = df.pivot_table(index=['Registration Number', 'Student Name'], columns='unit_code', values='score').reset_index()

# Fill NaN values with 0 if needed
df = df.fillna(0)

# Print the resulting DataFrame
print(df)
