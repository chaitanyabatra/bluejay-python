import pandas as pd
import datetime
from datetime import date, timedelta


def parse_date(date_str):
    try:
        # Try the original format
        return datetime.datetime.strptime(str(date_str), '%m/%d/%Y %I:%M %p')
    except ValueError:
        try:
            # Try an alternative format
            return datetime.datetime.strptime(str(date_str), '%Y-%m-%d %H:%M:%S')
        except ValueError:
            # If both formats fail, return None or handle it accordingly
            return None

def analyze_excel(file_path):
    df = pd.read_excel(file_path)

    employee_data = {}
    output=[]
    o1=[]
    o2=[]
    o3=[]

    for index, row in df.iterrows():
        position_id = row['Position ID']
        position_status = row['Position Status']
        time = parse_date(row['Time'])
        time_out = parse_date(row['Time Out'])
        timecard_hours = row['Timecard Hours (as Time)']
        employee_name = row['Employee Name']
        file_number = row['File Number']

        if position_status == 'Active' and pd.notna(time) and pd.notna(time_out):
            if employee_name not in employee_data:
                employee_data[employee_name] = {'position_id': position_id, 'shifts': []}

            employee_data[employee_name]['shifts'].append({
                'time': time,
                'time_out': time_out,
                'timecard_hours': timecard_hours
            })

    for employee_name, data in employee_data.items():
        shifts = data['shifts']
        flag1,flag2,flag3=False,False,False
        # Check for 7 consecutive days
        consecutive_days_count = 0
        for i in range(len(shifts) - 6):
            consecutive_dates = [shifts[i + j]['time'].date() for j in range(7)]
            #print(consecutive_dates) to check if needed
            sorted_dates = sorted(set(consecutive_dates))
            
            # Iterate through the sorted list and check for consecutive dates
            for i in range(len(sorted_dates)-6):
                #print(sorted_dates)
                if all(sorted_dates[i + j + 1] == sorted_dates[i + j] + timedelta(days= 1) for j in range(6)):
                    consecutive_days_count += 1

        # Check for less than 10 hours between shifts but more than 1 hour
        for i in range(len(shifts) - 1):
            time_between_shifts = (shifts[i + 1]['time'] - shifts[i]['time_out']).seconds // 3600
            if 1 < time_between_shifts < 10:
                #print(f"Employee {employee_name} (Position ID: {data['position_id']}) has less than 10 hours between shifts but more than 1 hour on {shifts[i]['time'].date()}")
                o1.append([employee_name,data['position_id']])
                flag2=True

        # Check for more than 14 hours in a single shift
        for shift in shifts:
            hours_worked = (shift['time_out'] - shift['time']).seconds // 3600
            if hours_worked > 14:
                #print(f"Employee {employee_name} (Position ID: {data['position_id']}) worked for more than 14 hours on {shift['time'].date()}")
                o2.append([employee_name,data['position_id']])
                flag3=True

        # Print information if the person has worked 7 consecutive days
        if consecutive_days_count > 0:
            #print(f"Employee {employee_name} (Position ID: {data['position_id']}) has worked 7 consecutive days {consecutive_days_count} times.")
            o3.append([employee_name,data['position_id']])
            flag1=True
        if (flag1 and flag2 and flag3):
            output.append([employee_name,data['position_id']])
    print("Employees that satisfy all three conditions are : ",output,"\n","\n")
    print("Employees that have less than 10 hours but more than 1 hour are between shifts : ",o1,"\n","\n")
    print("Employees that worked for more than 14 hours are : ",o2,"\n","\n")
    print("Employees that worked 7 consecutive days are : ",o3,"\n","\n")

# Example usage
file_path = '/Users/chaitanyabatra/Downloads/Assignment_Timecard.xlsx'  # Replace with the actual path to your downloaded file
analyze_excel(file_path)
