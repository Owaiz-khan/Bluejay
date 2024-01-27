import pandas as pd
from datetime import datetime, timedelta

def analyze_employee_data(file_path):
    df = pd.read_excel(file_path)

    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    df = df.sort_values(by=['Employee Name', 'Time'])

    results = {'7_consecutive_days': [], 'less_than_10_hours_between_shifts': [], 'more_than_14_hours_single_shift': []}

    for employee, employee_data in df.groupby('Employee Name'):
        # Check for 7 consecutive days worked
        consecutive_days = (employee_data['Time'].diff().dt.days == 1).rolling(7, min_periods=1).sum()
        if any(consecutive_days >= 7):
            results['7_consecutive_days'].append(employee)

        # Check for less than 10 hours between shifts but greater than 1 hour
        time_diff_between_shifts = employee_data['Time'].diff().fillna(pd.Timedelta(seconds=0))
        less_than_10_hours = time_diff_between_shifts.between('1 hour', '10 hours')
        if any(less_than_10_hours):
            results['less_than_10_hours_between_shifts'].append(employee)

        # Check for more than 14 hours in a single shift
        more_than_14_hours = (employee_data['Time Out'] - employee_data['Time']).dt.total_seconds() > 14 * 3600
        if any(more_than_14_hours):
            results['more_than_14_hours_single_shift'].append(employee)

    print("Employees who have worked for 7 consecutive days:", results['7_consecutive_days'])
    print("Employees who have less than 10 hours between shifts but greater than 1 hour:", results['less_than_10_hours_between_shifts'])
    print("Employees who have worked for more than 14 hours in a single shift:", results['more_than_14_hours_single_shift'])

file_path = 'Assignment_Timecard.xlsx'
analyze_employee_data(file_path)
