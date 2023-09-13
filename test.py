import re

# Define a regular expression pattern to match all variations
internal_marks_cell_pattern = re.compile(r'(int\.?|internal)\s*examiner\s*marks\s*/?\s*100', re.IGNORECASE)

# Test the pattern with sample variations
test_strings = [
    "internal examiner marks /100",
    "INTERNAL EXAMINER MARKS  /100",
    "INT.  EXAMINER MARKS  /100",
]

for test_string in test_strings:
    if internal_marks_cell_pattern.search(test_string):
        print(f"Matched: {test_string}")
    else:
        print(f"Not Matched: {test_string}")
