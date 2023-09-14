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


import pandas as pd

# Sample DataFrame
data = {'Student Name': ['Alice', 'Bob', 'Charlie', 'Kim'],
        'Recommendation': ['SUPP = 1 UNIT', 'PASS', 'SUPP = 3 UNITS', 'SPECIAL = 2 UNITS']}
consolidated_data_df = pd.DataFrame(data)

# Custom function to check if the Recommendation contains 'SUPP' and does not contain 'SPECIAL' and units information
def has_supp_units(recommendation):
    return 'SUPP' in recommendation and 'SPECIAL' not in recommendation and any(word.isdigit() for word in recommendation.split())

# Filter the DataFrame using the custom function
supp_students_df = consolidated_data_df[consolidated_data_df['Recommendation'].apply(has_supp_units)]

# Print the filtered DataFrame
print(supp_students_df)
