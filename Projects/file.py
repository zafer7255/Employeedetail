import openpyxl
from datetime import datetime, timedelta

# Define a function to check if a date is consecutive
def is_consecutive(date_list):
    # Convert datetime objects to strings
    date_strings = [date.strftime('%m/%d/%Y %I:%M %p') if isinstance(date, datetime) else date for date in date_list]
    
    dates = [datetime.strptime(date, '%m/%d/%Y %I:%M %p') for date in date_strings if date.strip()]
    
    # Check if there are any valid dates
    if not dates:
        return False
    
    dates.sort()
    expected_next_date = dates[0]

    for date in dates:
        if date != expected_next_date:
            return False
        expected_next_date += timedelta(days=1)

    return True

# Define a function to calculate the time difference between two timestamps
def calculate_time_difference(timestamp1, timestamp2):
    # Check if either timestamp is empty, and return a large timedelta in such cases
    if not isinstance(timestamp1, datetime) or not isinstance(timestamp2, datetime):
        return timedelta(hours=24 * 365)  # A large timedelta representing a year
    
    time_difference = timestamp2 - timestamp1

    return time_difference

# Initialize data structures to store results
consecutive_work_days = []
short_time_between_shifts = []
long_single_shifts = []

# Initialize the previous_time_out variable
previous_time_out = None

# Open and read the Excel file (provide the correct file path)
excel_file_path = '/home/zafer/Desktop/Projects/employeedetail.xlsx'
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Loop through the rows in the Excel sheet and perform analysis
for row in sheet.iter_rows(min_row=2, values_only=True):  # Start from the second row assuming the first row contains headers
    # Check if there are enough values in the row
    if len(row) < 8:
        continue  # Skip rows without enough values

    # Unpack only the required values
    position_id, position_status, time_in, time_out, _, _, _, employee_name = row[:8]

    if is_consecutive([time_in, time_out]):
        consecutive_work_days.append(employee_name)

    if previous_time_out is not None and timedelta(hours=1) < calculate_time_difference(previous_time_out, time_in) < timedelta(hours=10):
        short_time_between_shifts.append(employee_name)

    if calculate_time_difference(time_in, time_out) > timedelta(hours=14):
        long_single_shifts.append(employee_name)

    previous_time_out = time_out

# Print the results
print("Employees with 7 consecutive work days:")
print(consecutive_work_days)

print("\nEmployees with short time between shifts:")
print(short_time_between_shifts)

print("\nEmployees with long single shifts:")
print(long_single_shifts)

# Write the results to an output file
output_file_path = '/home/zafer/Desktop/Projects/output.txt'
with open(output_file_path, 'w') as output_file:
    output_file.write("Employees with 7 consecutive work days:\n")
    output_file.write("\n".join(consecutive_work_days))
    output_file.write("\n\nEmployees with short time between shifts:\n")
    output_file.write("\n".join(short_time_between_shifts))
    output_file.write("\n\nEmployees with long single shifts:\n")
    output_file.write("\n".join(long_single_shifts))

