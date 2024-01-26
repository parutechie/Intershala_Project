import pandas as pd
from pandas import NaT, Timestamp
from datetime import datetime, timedelta

def parse_datetime(date_str, time_str):
    if pd.isna(date_str) or pd.isna(time_str) or not isinstance(date_str, Timestamp):
        return NaT
    
    # Extract date and time components
    date_component = date_str.date()
    time_component = time_str.time()
    
    return datetime.combine(date_component, time_component)

def check_consecutive_days(shifts):
    for i in range(len(shifts) - 6):
        if (shifts[i + 6]['start'] - shifts[i]['end']).days == 6:
            return True
    return False

def check_time_between_shifts(shifts):
    for i in range(len(shifts) - 1):
        time_between_shifts = shifts[i + 1]['start'] - shifts[i]['end']
        if timedelta(hours=1) < time_between_shifts < timedelta(hours=10):
            return True
    return False

def check_long_shift(shift):
    return (shift['end'] - shift['start']).total_seconds() / 3600 > 14

def analyze_shifts(shifts):
    unique_messages = set()

    for employee_id, employee_shifts in shifts.items():
        if check_consecutive_days(employee_shifts):
            unique_messages.add(f"Employee {employee_id} has worked for 7 consecutive days.")

        for shift in employee_shifts:
            if check_time_between_shifts(employee_shifts):
                unique_messages.add(f"Employee {employee_id} has less than 10 hours but greater than 1 hour between shifts.")

            if check_long_shift(shift):
                unique_messages.add(f"Employee {employee_id} has worked for more than 14 hours in a single shift.")
    
    return unique_messages

def main(input_file, output_file):
    shifts = {}

    # Read Excel file using pandas
    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
        employee_id = row['Employee Name']
        date_str = row['Pay Cycle Start Date']
        start_time_str = row['Time']
        end_time_str = row['Time Out']

        start_time = parse_datetime(date_str, start_time_str)
        end_time = parse_datetime(date_str, end_time_str)

        if employee_id not in shifts:
            shifts[employee_id] = []

        shifts[employee_id].append({'start': start_time, 'end': end_time})

    unique_output = analyze_shifts(shifts)

    # Print unique messages to console
    for line in unique_output:
        print(line)

    # Write unique output to file
    with open(output_file, "w") as file:
        for line in unique_output:
            file.write(line + "\n")

if __name__ == "__main__":
    input_file = "Assignment_Timecard.xlsx"
    output_file = "output.txt"
    main(input_file, output_file)
