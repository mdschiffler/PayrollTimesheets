import pandas as pd
import sys
from datetime import datetime, timedelta
import os
from xlsxwriter.utility import quote_sheetname

def process_timesheet(csv_file, output_excel):
    if not os.path.exists(csv_file):
        print(f"Input file not found: {csv_file}")
        sys.exit(1)

    # Load main timesheet CSV
    df = pd.read_csv(csv_file)
    df.columns = [col.strip() for col in df.columns]

    # Load rates CSV (ID,RATE,START,EXTRA,DETAILS) from parent folder of the input CSV
    csv_dir = os.path.dirname(csv_file)
    parent_dir = os.path.abspath(os.path.join(csv_dir, os.pardir))
    rates_file = os.path.join(parent_dir, "timesheet-rates.csv")
    if os.path.exists(rates_file):
        rates_df = pd.read_csv(rates_file)
        rates_df['START'] = pd.to_datetime(rates_df['START'], errors='coerce')
        rates_dict = {}
        for _, row in rates_df.iterrows():
            rates_dict[row['ID']] = {
                'RATE': row['RATE'],
                'START': row['START'],
                'EXTRA': row['EXTRA'],
                'DETAILS': row['DETAILS'] if 'DETAILS' in row else ''
            }
    else:
        print("Warning: 'timesheet-rates.csv' not found. All rates will default to $0.")
        rates_dict = {}

    # Combine Punch Date + Time, supporting multiple date formats
    datetime_str = df['Punch Date'].str.strip() + " " + df['Attendance record'].str.strip()

    # Try parsing with both possible date formats
    df['Datetime'] = pd.to_datetime(datetime_str, format='%Y-%m-%d %H:%M:%S', errors='coerce')
    mask_failed = df['Datetime'].isna()
    if mask_failed.any():
        df.loc[mask_failed, 'Datetime'] = pd.to_datetime(
            datetime_str[mask_failed], format='%m/%d/%Y %H:%M:%S', errors='coerce'
        )

    # Drop rows that still couldn't be parsed
    if df['Datetime'].isna().any():
        print("Warning: some datetime entries could not be parsed and will be skipped.")
        df = df.dropna(subset=['Datetime'])

    # Group by person/date
    grouped = df.groupby(['Person ID', 'Person Name', 'Punch Date'])
    records = []

    for (person_id, person_name, punch_date), group in grouped:
        group_sorted = group.sort_values(by='Datetime')
        check_in = group_sorted.iloc[0]['Datetime']
        check_out = group_sorted.iloc[-1]['Datetime']
        hours_worked = round((check_out - check_in).total_seconds() / 3600, 2)
        records.append({
            'Person ID': person_id,
            'Person Name': person_name,
            'Location': 'Maru',
            'Date': punch_date,
            'Start': check_in.strftime("%H:%M:%S"),
            'End': check_out.strftime("%H:%M:%S"),
            'Hours': hours_worked,
            'Details': ''  # placeholder for new column
        })

    # Organize records per person
    records_df = pd.DataFrame(records)
    persons = {}
    for (person_id, person_name), group in records_df.groupby(['Person ID', 'Person Name']):
        df_person = group.sort_values(by='Date').reset_index(drop=True)
        # Insert Details column after Hours
        cols = list(df_person.columns)
        if 'Details' in cols:
            cols.remove('Details')
            idx = cols.index('Hours') + 1
            cols.insert(idx, 'Details')
            df_person = df_person[cols]
        # Reorder columns to have Location first, then Date, then others
        cols = ['Location', 'Date'] + [col for col in df_person.columns if col not in ['Location', 'Date', 'Person ID', 'Person Name']]
        df_person = df_person[cols]
        persons[(person_id, person_name)] = df_person

    # Write to Excel
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        workbook = writer.book
        # Common formats
        header_format = workbook.add_format({'bold': True, 'border': 1})
        currency_format = workbook.add_format({'num_format': '$#,##0.00'})
        red_bg_format = workbook.add_format({'bg_color': '#FF0000', 'num_format': '$#,##0.00'})
        soft_red_format = workbook.add_format({'bg_color': '#FF9999', 'num_format': '$#,##0.00'})
        light_green_text_format = workbook.add_format({'bg_color': '#CCFFCC'})
        light_green_currency_format = workbook.add_format({'bg_color': '#CCFFCC', 'num_format': '$#,##0.00'})
        summary_header_format = workbook.add_format({'bold': True, 'border': 1})

        # Summary sheet created first so it is the first tab
        summary_sheet = workbook.add_worksheet("Summary")
        writer.sheets["Summary"] = summary_sheet
        summary_entries = []

        for (person_id, person_name), df_person in persons.items():
            sheet_name = f"{person_id} - {person_name}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            start_row = 3
            df_person.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)

            worksheet = writer.sheets[sheet_name]

            # Header and formatting
            # Write person info row with light green background
            worksheet.write(0, 0, f"Person ID: {person_id}, Name: {person_name}", light_green_text_format)
            worksheet.write(0, 1, '', light_green_text_format)
            # Insert row for timesheet period (parsed from filename)
            import re
            match = re.search(r'(\d{2}-\d{2}-\d{4})', os.path.basename(csv_file))
            if match:
                end_date = datetime.strptime(match.group(1), '%m-%d-%Y')
                start_date = end_date - timedelta(days=6)
                worksheet.write(1, 0, f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                worksheet.write_blank(2, 0, None)
            worksheet.set_column('A:D', 20)
            worksheet.set_column('E:E', 30)  # for Details column

            # Row indices
            n = len(df_person)
            total_row_idx        = start_row + n + 1
            rate_row_idx         = total_row_idx + 1

            def write_location_section(start_row, header_title, placeholders=None):
                placeholders = placeholders or ["Apt X", "Apt X"]
                worksheet.write(start_row, 0, header_title, header_format)
                section_headers = ['Date', 'Check-in', 'Check-out', 'Rate $', 'Details']
                for col, name in enumerate(section_headers, start=1):
                    worksheet.write(start_row, col, name, header_format)

                data_start = start_row + 1
                for offset, text in enumerate(placeholders):
                    worksheet.write(data_start + offset, 0, text)

                total_row = data_start + len(placeholders)
                first_excel_row = data_start + 1
                last_excel_row = data_start + len(placeholders)
                worksheet.write(total_row, 0, 'Total $')
                worksheet.write_formula(total_row, 4, f"=SUM(E{first_excel_row}:E{last_excel_row})", currency_format)
                return {
                    "total_row": total_row,
                    "next_start": total_row + 2,
                    "data_start_excel": first_excel_row,
                    "data_end_excel": last_excel_row,
                }

            current_section_row = rate_row_idx + 4
            section_totals_excel_rows = []
            section_definitions = [
                ("Mango Villas", ["Apt X", "Apt X"]),
                ("Casa Damisela", ["Apt X", "Apt X"]),
                ("Other", ["Details here", ""])
            ]
            section_clean_ranges = []
            for section_header, placeholders in section_definitions:
                section_info = write_location_section(
                    current_section_row, section_header, placeholders
                )
                section_totals_excel_rows.append(section_info["total_row"] + 1)
                section_clean_ranges.append(
                    (section_info["data_start_excel"], section_info["data_end_excel"])
                )
                current_section_row = section_info["next_start"]

            # Summary header directly after the location sections
            summary_header_row = current_section_row
            worksheet.write(summary_header_row, 0, "Summary", summary_header_format)
            worksheet.merge_range(summary_header_row, 0, summary_header_row, 5, "Summary", summary_header_format)

            extras_row_idx = summary_header_row + 1
            subtotal_row_idx = extras_row_idx + 1
            worksheet.write(subtotal_row_idx, 0, "Subtotal $")

            # Define Excel row numbers (1-based)
            rate_excel_row = rate_row_idx + 1
            total_excel_row = total_row_idx + 1
            extras_excel_row = extras_row_idx + 1

            # Add new Total $ row immediately below Rate $
            total_rate_row_idx = rate_row_idx + 1
            worksheet.write(rate_row_idx, 0, "Rate $")
            worksheet.write_number(rate_row_idx, 4, 0, currency_format)  # placeholder, will overwrite below
            worksheet.write(total_rate_row_idx, 0, "Total $")
            worksheet.write_formula(total_rate_row_idx, 4, f"=E{rate_row_idx+1} * E{total_row_idx+1}", currency_format)

            # Total hours
            total_hours = round(df_person['Hours'].sum(), 2)
            worksheet.write(total_row_idx, 0, "Total hours")
            worksheet.write(total_row_idx, 4, total_hours)

            # Rate (from lookup)
            rate = 0
            start_date = None
            extra_val = 0
            details_val = ''
            if person_id in rates_dict:
                rate = float(rates_dict[person_id]['RATE'])
                start_date = rates_dict[person_id]['START']
                extra_val = rates_dict[person_id]['EXTRA']
                details_val = str(rates_dict[person_id]['DETAILS']) if pd.notna(rates_dict[person_id]['DETAILS']) else ""

            worksheet.write_number(rate_row_idx, 4, rate, currency_format)

            # Extras and Annual withheld exclusion
            worksheet.write(extras_row_idx, 0, "Extras $")

            today = datetime.now()
            show_red = False
            if start_date is not None:
                if (today - start_date).days < 28 or today.month == 1:
                    # Within 4 weeks or January
                    extras_amount = 0
                    exclusion_amount = 0
                    show_red = True
                else:
                    extras_amount = extra_val
                    exclusion_amount = 500
            else:
                # No start date info, default to extras_val and 500
                extras_amount = extra_val
                exclusion_amount = 500

            worksheet.write_number(extras_row_idx, 4, extras_amount, currency_format)

            # Write DETAILS (from timesheet-rates.csv) in the next column (F)
            if details_val:
                worksheet.write(extras_row_idx, 5, details_val)

            exclusion_row_idx = extras_row_idx + 2
            worksheet.write(exclusion_row_idx, 0, "Annual withheld, incl. today $ (limit $500)")
            if show_red:
                worksheet.write_number(exclusion_row_idx, 4, exclusion_amount, soft_red_format)
            else:
                worksheet.write_number(exclusion_row_idx, 4, exclusion_amount, currency_format)

            # Define withheld_row_idx before using it
            withheld_row_idx = exclusion_row_idx + 1
            total_dollar_idx = withheld_row_idx + 2
            # Amount withheld (10% of Subtotal if Annual withheld exclusion = 500)
            worksheet.write(withheld_row_idx, 0, "10% withheld today $")
            exclusion_excel_row = exclusion_row_idx + 1
            withheld_excel_row = withheld_row_idx + 1
            # If exclusion cell equals 500 then withhold 10% of subtotal, else 0
            withheld_formula = f"=IF(E{exclusion_excel_row}=500,ROUNDDOWN(E{subtotal_row_idx+1}*0.10,2),0)"
            worksheet.write_formula(withheld_row_idx, 4, withheld_formula, currency_format)

            # Total $ (calculated)
            worksheet.write(total_dollar_idx, 0, "Total $")
            # Final Total = Subtotal - Amount withheld today
            total_excel_row = subtotal_row_idx + 1
            withheld_excel_row = withheld_row_idx + 1
            final_total_formula = f"=E{total_excel_row} - E{withheld_excel_row}"
            worksheet.write_formula(total_dollar_idx, 4, final_total_formula, light_green_currency_format)

            # Update subtotal formula to reference new Total $ row
            section_total_refs = [f"E{row}" for row in section_totals_excel_rows]
            subtotal_parts = [f"E{total_rate_row_idx+1}", *section_total_refs, f"E{extras_excel_row}"]
            subtotal_formula = "=" + " + ".join(subtotal_parts)
            worksheet.write_formula(subtotal_row_idx, 4, subtotal_formula, currency_format)

            # Capture summary references
            hours_cell = f"={quote_sheetname(sheet_name)}!E{total_row_idx + 1}"
            total_cell = f"={quote_sheetname(sheet_name)}!E{total_dollar_idx + 1}"

            sheet_ref = quote_sheetname(sheet_name)
            if n > 0:
                main_start_excel = start_row + 2
                main_end_excel = start_row + n + 1
                main_clean_count = f"COUNTA({sheet_ref}!B{main_start_excel}:B{main_end_excel})"
            else:
                main_clean_count = "0"

            section_clean_counts = [
                f'COUNTIF({sheet_ref}!E{start}:E{end},"<>")' for start, end in section_clean_ranges
            ]
            cleans_parts = [main_clean_count, *section_clean_counts]
            cleans_formula = "=" + "+".join(cleans_parts) if cleans_parts else "=0"

            summary_entries.append((sheet_name, hours_cell, cleans_formula, total_cell))

        # Populate summary sheet
        summary_sheet.write(0, 0, "Person", summary_header_format)
        summary_sheet.write(0, 1, "Total Hours", summary_header_format)
        summary_sheet.write(0, 2, "Total Cleans", summary_header_format)
        summary_sheet.write(0, 3, "Total $", summary_header_format)
        for idx, (label, hours_ref, cleans_ref, total_ref) in enumerate(summary_entries, start=1):
            summary_sheet.write(idx, 0, label)
            summary_sheet.write_formula(idx, 1, hours_ref)
            summary_sheet.write_formula(idx, 2, cleans_ref)
            summary_sheet.write_formula(idx, 3, total_ref, currency_format)
        if summary_entries:
            total_row = len(summary_entries) + 1
            summary_sheet.write(total_row, 0, "All sheets total", summary_header_format)
            summary_sheet.write_formula(total_row, 1, f"=SUM(B2:B{total_row})")
            summary_sheet.write_formula(total_row, 2, f"=SUM(C2:C{total_row})")
            summary_sheet.write_formula(total_row, 3, f"=SUM(D2:D{total_row})", light_green_currency_format)
        summary_sheet.set_column('A:A', 35)
        summary_sheet.set_column('B:D', 18)

    print(f"Excel file '{output_excel}' created successfully.")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python export_timesheet.py <input_csv> <output_excel>")
        sys.exit(1)

    input_csv = sys.argv[1]
    output_excel = sys.argv[2]
    process_timesheet(input_csv, output_excel)
