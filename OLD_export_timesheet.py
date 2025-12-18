import pandas as pd
import sys
from datetime import datetime

def process_timesheet(csv_file, output_excel):
    # Read CSV file
    df = pd.read_csv(csv_file)
    
    # Clean column names (strip extra spaces)
    df.columns = [col.strip() for col in df.columns]
    
    # Combine Punch Date and Attendance record into a datetime column.
    # Assuming "Attendance record" is in HH:MM:SS format and "Punch Date" is YYYY-MM-DD.
    df['Datetime'] = pd.to_datetime(
        df['Punch Date'].str.strip() + " " + df['Attendance record'].str.strip(), 
        format='%Y-%m-%d %H:%M:%S'
    )
    
    # Group by Person ID, Person Name, and Punch Date to get check‑in and check‑out per day.
    grouped = df.groupby(['Person ID', 'Person Name', 'Punch Date'])
    
    records = []
    for (person_id, person_name, punch_date), group in grouped:
        # Assume the earliest record is check‑in and the latest is check‑out.
        group_sorted = group.sort_values(by='Datetime')
        check_in = group_sorted.iloc[0]['Datetime']
        check_out = group_sorted.iloc[-1]['Datetime']
        hours_worked = round((check_out - check_in).total_seconds() / 3600, 2)
        records.append({
            'Person ID': person_id,
            'Person Name': person_name,
            'Punch Date': punch_date,
            'Check-in': check_in.strftime("%H:%M:%S"),
            'Check-out': check_out.strftime("%H:%M:%S"),
            'Hours Worked': hours_worked
        })
    
    # Convert the list of records into a DataFrame.
    records_df = pd.DataFrame(records)
    
    # Create a dictionary to hold each person's DataFrame.
    persons = {}
    for (person_id, person_name), group in records_df.groupby(['Person ID', 'Person Name']):
        # Sort by Punch Date.
        df_person = group.sort_values(by='Punch Date').reset_index(drop=True)
        persons[(person_id, person_name)] = df_person.drop(columns=['Person ID', 'Person Name'])
    
    # Write the data to an Excel file with one sheet per person.
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        for (person_id, person_name), df_person in persons.items():
            # Construct sheet name.
            sheet_name = f"{person_id} - {person_name}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            
            # Set the starting row for the table.
            # Row 0: custom header with person's info.
            # Row 1: blank.
            # Row 2: table header (written by to_excel) and data starts at row 3.
            start_row = 2
            df_person.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
            
            # Get the workbook and worksheet objects.
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Write the custom header (person's info) in the first row.
            worksheet.write(0, 0, f"Person ID: {person_id}, Name: {person_name}")
            # Row 1 is left blank.
            
            # Adjust column widths so data is not cut off.
            worksheet.set_column('A:D', 20)
            
            # The table written by to_excel includes a header row plus the data rows.
            # Let n be the number of data rows (excluding the header).
            n = len(df_person)
            # The table occupies rows: header at row index = start_row (i.e. row 2 in Excel) 
            # and data rows from row index start_row+1 to start_row+n.
            # The extra rows will be appended after that.
            total_row_idx    = start_row + n + 1  # Total row.
            empty_row_idx    = total_row_idx + 1    # Empty row.
            rate_row_idx     = empty_row_idx + 1    # Rate $ row.
            extras_row_idx   = rate_row_idx + 1       # Extras row.
            total_dollar_idx = extras_row_idx + 1     # Total $ row.
            
            # Calculate total hours from the "Hours Worked" column.
            total_hours = round(df_person['Hours Worked'].sum(), 2)
            
            # Write the Total row.
            # Column mapping: A: Punch Date, B: Check-in, C: Check-out, D: Hours Worked.
            worksheet.write(total_row_idx, 0, "Total")
            worksheet.write(total_row_idx, 3, total_hours)
            
            # Leave the empty row (empty_row_idx) blank.
            
            # Write the Rate $ row with a placeholder value of 0.
            worksheet.write(rate_row_idx, 0, "Rate $")
            worksheet.write(rate_row_idx, 3, 0)
            
            # Write the Extras row with a placeholder value of 0.
            worksheet.write(extras_row_idx, 0, "Extras $")
            worksheet.write(extras_row_idx, 3, 0)
            
            # Write the Total $ row with a formula.
            worksheet.write(total_dollar_idx, 0, "Total $")
            # Convert our 0-indexed row numbers to Excel's 1-indexed row numbers.
            total_excel_row  = total_row_idx + 1   # For the Total row.
            rate_excel_row   = rate_row_idx + 1      # For the Rate $ row.
            extras_excel_row = extras_row_idx + 1    # For the Extras row.
            # The formula calculates: (Rate $ * Total Hours) + Extras,
            # then rounds up to the nearest dollar and adds 1 cent.
            formula = f"=CEILING((D{rate_excel_row} * D{total_excel_row}) + D{extras_excel_row}, 1) + 0.01"
            worksheet.write_formula(total_dollar_idx, 3, formula)
    
    print(f"Excel file '{output_excel}' created successfully.")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python export_timesheet.py <input_csv> <output_excel>")
        sys.exit(1)
    
    input_csv = sys.argv[1]
    output_excel = sys.argv[2]
    process_timesheet(input_csv, output_excel)
