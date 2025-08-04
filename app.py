from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Format time with AM/PM
def format_time_with_ampm(time_val):
    if pd.isna(time_val) or time_val == '':
        return ''
    try:
        time_obj = pd.to_datetime(str(time_val)).time()
        am_pm = 'AM' if time_obj.hour < 12 else 'PM'
        return f"{time_obj.strftime('%H:%M')} {am_pm}"
    except:
        return str(time_val)

# Update all sheets in Excel and save as full workbook
def update_attendance_file(input_file, output_file):
    excel_data = pd.read_excel(input_file, sheet_name=None, skiprows=1)  # Load all sheets after skipping metadata
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    updated_sheet_count = 0

    for sheet_name, df in excel_data.items():
        if df.empty or df.shape[1] < 4:
            print(f"Skipping sheet '{sheet_name}' due to insufficient data.")
            continue

        try:
            df.columns = ['Date', 'Day', 'First Check In', 'Last Check Out']
            df = df.dropna(subset=['Date'])
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])

            df['Date'] = df['Date'].dt.normalize()
            df['Day'] = df['Date'].dt.day_name()

            start_date = df['Date'].min().replace(day=1)
            end_date = (start_date + pd.offsets.MonthEnd(0)).normalize()
            all_dates = pd.date_range(start=start_date, end=end_date)

            new_df = pd.DataFrame({
                'Date': all_dates,
                'Day': all_dates.day_name(),
                'First Check In': '',
                'Last Check Out': ''
            })

            merged_df = pd.merge(
                new_df,
                df[['Date', 'First Check In', 'Last Check Out']],
                on='Date',
                how='left',
                suffixes=('', '_existing')
            )

            merged_df['First Check In'] = merged_df['First Check In_existing'].combine_first(merged_df['First Check In'])
            merged_df['Last Check Out'] = merged_df['Last Check Out_existing'].combine_first(merged_df['Last Check Out'])

            merged_df['First Check In'] = merged_df['First Check In'].apply(format_time_with_ampm)
            merged_df['Last Check Out'] = merged_df['Last Check Out'].apply(format_time_with_ampm)

            merged_df['Date'] = merged_df['Date'].dt.strftime('%Y-%m-%d')
            merged_df = merged_df[['Date', 'Day', 'First Check In', 'Last Check Out']]

            # Write updated sheet to output Excel
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            updated_sheet_count += 1

        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")
            continue

    if updated_sheet_count == 0:
        raise ValueError("No valid sheets found to update.")

    writer.save()

# Web interface
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file.filename.endswith('.xlsx'):
            input_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            output_path = os.path.join(PROCESSED_FOLDER, 'updated_' + uploaded_file.filename)
            try:
                uploaded_file.save(input_path)
                update_attendance_file(input_path, output_path)
                return send_file(output_path, as_attachment=True)
            except Exception as e:
                return f"Error processing file: {str(e)}"
        else:
            return "Please upload a valid Excel (.xlsx) file."
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
