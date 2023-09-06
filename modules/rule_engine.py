# Make rules

import pandas as pd 

# Define a function to calculate the grade
def calculate_grade(mean):
    if pd.isna(mean):
        return ''
    if mean >= 70:
        return 'A'
    elif 60 <= mean < 70:
        return 'B'
    elif 50 <= mean < 60:
        return 'C'
    elif 40 <= mean < 50:
        return 'D'
    else:
        return 'E'
    

