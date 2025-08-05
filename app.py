from flask import Flask, render_template, request, send_file, session, redirect, url_for, flash
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = '1234'  # Use a secure random key in production

UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def format_time_with_ampm(time_val):
    if pd.isna(time_val) or time_val == '':
        return ''
    try:
        time_obj = pd.to_datetime(str(time_val)).time()
        am_pm = 'AM' if time_obj.hour < 12 else 'PM'
        return f"{time_obj.strftime('%H:%M')} {am_pm}"
    except:
        return str(time_val)

def update_attendance_file(input_file, output_file):
    excel_data = pd.read_excel(input_file, sheet_name=None, skiprows=1)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    updated_sheet_count = 0

    for sheet_name, df in excel_data.items():
        if df.empty or df.shape[1] < 4:
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

            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            updated_sheet_count += 1
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")
            continue

    if updated_sheet_count == 0:
        raise ValueError("No valid sheets found to update.")
    writer.save()

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == 'admin' and password == 'thanuja@0420':
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('Invalid credentials', 'error')
            return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file and uploaded_file.filename.endswith('.xlsx'):
            input_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            output_path = os.path.join(PROCESSED_FOLDER, 'updated_' + uploaded_file.filename)
            try:
                uploaded_file.save(input_path)
                update_attendance_file(input_path, output_path)
                return send_file(output_path, as_attachment=True)
            except Exception as e:
                flash(f"Error processing file: {str(e)}", 'error')
                return redirect(url_for('index'))
        else:
            flash("Please upload a valid Excel (.xlsx) file.", 'error')
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
