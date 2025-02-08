import pandas as pd
from datetime import datetime, timedelta

# Input the first Monday date of the semester
first_monday = input("Enter the first day of the semester (dd/mm/yyyy): ")
first_monday = datetime.strptime(first_monday, "%d/%m/%Y")

# Read the Excel file
df = pd.read_excel('time-table.xlsx')

# Delete the first row
df = df.drop(index=0)

# Split the "Thời gian" column into "Start Time" and "End Time"
df[['Start Time', 'End Time']] = df['11'].str.split('-', expand=True)

# Convert "Start Time" and "End Time" to datetime objects (time only)
df['Start Time'] = pd.to_datetime(df['Start Time'], format='%H%M').dt.strftime('%H:%M')
df['End Time'] = pd.to_datetime(df['End Time'], format='%H%M').dt.strftime('%H:%M')

df = df.rename(columns={'6': 'Subject'})
df['Start Date'] = '10/02/2025'

# Rename the column "Thứ" to "Day of Week"
df = df.rename(columns={'10': 'Day of Week'})
df = df.rename(columns={'15': 'Week'})

# Mapping dictionary for day numbers to weekday names (assuming 2 is Monday, 3 is Tuesday, etc.)
weekday_mapping = {
    2: 'Monday',
    3: 'Tuesday',
    4: 'Wednesday',
    5: 'Thursday',
    6: 'Friday',
    7: 'Saturday',
    8: 'Sunday' 
}

# Convert numbers in "Day of Week" column to weekday names
df['Day of Week'] = df['Day of Week'].map(weekday_mapping)

# Function to calculate the correct class date based on the week number and the day of the week
def calculate_class_dates(weeks_str, day_of_week, first_monday):
    dates = []
    weeks = weeks_str.split(',')
    
    # Find the weekday number (e.g., Monday = 0, Sunday = 6)
    weekday_num = (list(weekday_mapping.values()).index(day_of_week))

    for week in weeks:
        if '-' in week:
            # Range of weeks
            start_week, end_week = map(int, week.split('-'))
            for week_num in range(start_week, end_week + 1):
                # Calculate the start date for this week
                start_of_week = first_monday + timedelta(weeks=week_num - 24)
                # Adjust for the correct day of the week
                days_diff = (weekday_num - start_of_week.weekday()) % 7
                class_date = start_of_week + timedelta(days=days_diff)
                dates.append(class_date)
        else:
            # Single week
            week_num = int(week)
            start_of_week = first_monday + timedelta(weeks=week_num - 24)
            days_diff = (weekday_num - start_of_week.weekday()) % 7
            class_date = start_of_week + timedelta(days=days_diff)
            dates.append(class_date)
    
    return dates

# Calculate the class dates for each row
df['Class Dates'] = df.apply(lambda row: calculate_class_dates(row['Week'], row['Day of Week'], first_monday), axis=1)

# Function to format the event information and add to CSV
def generate_event_rows(row):
    event_rows = []
    for class_date in row['Class Dates']:
        start_datetime = f"{class_date.strftime('%d/%m/%Y')} {row['Start Time']}"
        end_datetime = f"{class_date.strftime('%d/%m/%Y')} {row['End Time']}"
        
        start_time = datetime.strptime(start_datetime, '%d/%m/%Y %H:%M')
        end_time = datetime.strptime(end_datetime, '%d/%m/%Y %H:%M')
        
        event_row = {
            'Subject': row['Subject'],
            'StartDate': start_time.strftime('%d/%m/%Y'),
            'EndDate': start_time.strftime('%d/%m/%Y'),  # Ensure StartDate = EndDate
            'Start Time': start_time.strftime('%H:%M'),
            'End Time': end_time.strftime('%H:%M'),
            'All Day Event': False,
            'Location': row['16'],  # Assuming column 16 contains location data
            'Private': False
        }
        event_rows.append(event_row)
    
    return event_rows

# Generate event rows for all classes
event_data = []
for index, row in df.iterrows():
    event_data.extend(generate_event_rows(row))

# Convert the event data into a DataFrame
event_df = pd.DataFrame(event_data)

# Save the event data to a new CSV file
event_df.to_csv('class_schedule_events.csv', index=False)

print("Event schedule saved to class_schedule_events.csv")
