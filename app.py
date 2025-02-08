import os
from flask import Flask, request, render_template, send_from_directory, after_this_request
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# Ensure the uploads folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Function to calculate class dates (same as in your code)
def calculate_class_dates(weeks_str, day_of_week, first_monday):
    dates = []
    weeks = weeks_str.split(',')
    weekday_mapping = {
        'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6
    }
    weekday_num = weekday_mapping[day_of_week]

    for week in weeks:
        if '-' in week:
            # Range of weeks
            start_week, end_week = map(int, week.split('-'))
            for week_num in range(start_week, end_week + 1):
                start_of_week = first_monday + timedelta(weeks=week_num - 24)
                days_diff = (weekday_num - start_of_week.weekday()) % 7
                class_date = start_of_week + timedelta(days=days_diff)
                dates.append(class_date)
        else:
            week_num = int(week)
            start_of_week = first_monday + timedelta(weeks=week_num - 24)
            days_diff = (weekday_num - start_of_week.weekday()) % 7
            class_date = start_of_week + timedelta(days=days_diff)
            dates.append(class_date)
    
    return dates

# Main page route
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and class schedule generation
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return "Invalid file format. Please upload an .xlsx file."
    
    # Save the uploaded file
    filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filename)
    
    # Get the first Monday from the form
    first_monday_str = request.form['first_monday']
    first_monday = datetime.strptime(first_monday_str, "%d/%m/%Y")
    
    # Read the uploaded Excel file
    df = pd.read_excel(filename)
    df = df.drop(index=0)
    df[['Start Time', 'End Time']] = df['11'].str.split('-', expand=True)
    df['Start Time'] = pd.to_datetime(df['Start Time'], format='%H%M').dt.strftime('%H:%M')
    df['End Time'] = pd.to_datetime(df['End Time'], format='%H%M').dt.strftime('%H:%M')
    df = df.rename(columns={'6': 'Subject', '10': 'Day of Week', '15': 'Week'})
    df['Day of Week'] = df['Day of Week'].map({
        2: 'Monday', 3: 'Tuesday', 4: 'Wednesday', 5: 'Thursday', 6: 'Friday', 7: 'Saturday', 8: 'Sunday'
    })

    # Calculate the class dates
    df['Class Dates'] = df.apply(lambda row: calculate_class_dates(row['Week'], row['Day of Week'], first_monday), axis=1)

    # Generate event data for CSV output
    event_data = []
    for _, row in df.iterrows():
        for class_date in row['Class Dates']:
            start_datetime = f"{class_date.strftime('%d/%m/%Y')} {row['Start Time']}"
            end_datetime = f"{class_date.strftime('%d/%m/%Y')} {row['End Time']}"
            start_time = datetime.strptime(start_datetime, '%d/%m/%Y %H:%M')
            end_time = datetime.strptime(end_datetime, '%d/%m/%Y %H:%M')
            
            event_data.append({
                'Subject': row['Subject'],
                'StartDate': start_time.strftime('%d/%m/%Y'),
                'EndDate': start_time.strftime('%d/%m/%Y'),
                'Start Time': start_time.strftime('%H:%M'),
                'End Time': end_time.strftime('%H:%M'),
                'All Day Event': False,
                'Location': row['16'],
                'Private': False
            })

    # Save the event data to CSV
    output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'class_schedule_events.csv')
    event_df = pd.DataFrame(event_data)
    event_df.to_csv(output_file, index=False)

    # Delete the uploaded file to keep the server static
    os.remove(filename)

    # Return the CSV file URL for downloading
    return render_template('index.html', file_url=f"/uploads/class_schedule_events.csv")

# Route to serve uploaded files and delete them after download
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    # Delete the file after sending it to the user
    @after_this_request
    def delete_file(response):
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file: {e}")
        return response
    
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == "__main__":
    app.run(debug=True)
