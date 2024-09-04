import pandas as pd
import datetime
import streamlit as st
from io import BytesIO

def load_data(file):
    return pd.read_excel(file, sheet_name=None)  # Load all sheets

def parse_datetime(date_str, time_str):
    return datetime.datetime.strptime(f"{date_str} {time_str}", "%d/%m/%Y %H:%M:%S")

def filter_data_for_day(data, shift_date):
    shift_start_datetime = parse_datetime(shift_date, '00:00:00')  # Start of the day
    shift_end_datetime = parse_datetime(shift_date, '23:59:59')  # End of the day

    data['DateTime'] = pd.to_datetime(data['Date'] + ' ' + data['Punch Time'], format="%d/%m/%Y %H:%M:%S")
   
    # Filter data for the given shift date
    filtered_data = data[
        (data['DateTime'] >= shift_start_datetime) &
        (data['DateTime'] <= shift_end_datetime)
    ]
   
    return filtered_data

def calculate_morning_shift(data):
    total_login_logout_time = datetime.timedelta()
    total_break_time = datetime.timedelta()
 
    first_login = None
    last_logout = None
    in_time = None
    prev_out_time = None

    for _, row in data.iterrows():
        current_datetime = row['DateTime']
 
        if row['I/O Type'] == 'IN':
            if first_login is None:
                first_login = current_datetime
            in_time = current_datetime
            if prev_out_time and prev_out_time < in_time:
                total_break_time += in_time - prev_out_time
        elif row['I/O Type'] == 'OUT' and in_time:
            last_logout = current_datetime
            prev_out_time = current_datetime
 
    if first_login and last_logout:
        total_login_logout_time = last_logout - first_login
 
    def timedelta_to_hours_minutes(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours} hours, {minutes} minutes"
 
    total_hours_worked = total_login_logout_time - total_break_time

    results = {
        'Date': data.iloc[0]['Date'],
        'Name': data.iloc[0]['Name'],
        'Total Time from Login to Logout (including breaks)': timedelta_to_hours_minutes(total_login_logout_time),
        'Break Time': timedelta_to_hours_minutes(total_break_time),
        'Total Hours Worked (excluding breaks)': timedelta_to_hours_minutes(total_hours_worked)
    }
    return results
 
def process_all_sheets(file):
    sheets = pd.read_excel(file, sheet_name=None)
    results_dict = {}
   
    for sheet_name, data in sheets.items():
        st.write(f"Processing sheet: {sheet_name}")
       
        results_data = []
       
        # Process data user-wise and date-wise in the same order as the input
        for name, user_data in data.groupby('Name', sort=False):
            working_dates = user_data['Date'].unique()
            for shift_date in working_dates:
                filtered_data = filter_data_for_day(user_data, shift_date)
                if not filtered_data.empty:
                    result = calculate_morning_shift(filtered_data)
                    if result:
                        results_data.append(result)
        
        if results_data:
            results_df = pd.DataFrame(results_data)
            # Save to Excel
            with BytesIO() as buffer:
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    results_df.to_excel(writer, sheet_name=f'Results_{sheet_name}', index=False)
               
                buffer.seek(0)
                results_dict[sheet_name] = buffer.getvalue()
   
    return results_dict
 
def main():
    st.title("Morning Shift Hours Calculator")
 
    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
 
    if uploaded_file is not None:
        if st.button('Process File'):
            try:
                # Process all sheets in the uploaded file and get the new Excel files as bytes
                results_dict = process_all_sheets(uploaded_file)
 
                for sheet_name, new_file_bytes in results_dict.items():
                    st.write(f"Results for sheet: {sheet_name}")
                    with pd.ExcelFile(BytesIO(new_file_bytes)) as xls:
                        df = pd.read_excel(xls, sheet_name=f'Results_{sheet_name}')
                        st.dataframe(df)
                       
                    st.download_button(
                        label=f"Download Processed File for {sheet_name}",
                        data=new_file_bytes,
                        file_name=f"Processed_Shift_Data_{sheet_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
           
            except Exception as e:
                st.error(f"Error: {e}")
 
if __name__ == "__main__":
    main()
