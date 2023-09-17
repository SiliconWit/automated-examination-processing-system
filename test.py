import json

# Specify the path to your JSON file
json_file_path = "data/units/mechatronics_engineering_units.json"

# Load your JSON data from the file
with open(json_file_path, "r") as json_file:
    data = json.load(json_file)

# Extract unit codes from the 4th Year 1st Semester and 4th Year 2nd Semester
fourth_year_first_semester = data.get("4th Year", {}).get("1st Semester", [])
fourth_year_second_semester = data.get("4th Year", {}).get("2nd Semester", [])

# Extract unit codes from the units
unit_codes_first_semester = [unit["Unit Code"] for unit in fourth_year_first_semester]
unit_codes_second_semester = [unit["Unit Code"] for unit in fourth_year_second_semester]

# # Create your semester order lists
# semester_order_lists = {
#     "4th_yr_1st_semester": unit_codes_first_semester,
#     "4th_yr_2nd_semester": unit_codes_second_semester,
# }

# # Print the order lists (you can save them to a TOML file if needed)
# for semester, unit_codes in semester_order_lists.items():
#     print(f"{semester} = {unit_codes}")

print(unit_codes_first_semester)
print(unit_codes_second_semester)