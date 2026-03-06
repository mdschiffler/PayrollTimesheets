import pandas as pd
import sys
from datetime import datetime, timedelta
import os
import re
import unicodedata
from xlsxwriter.utility import quote_sheetname

def normalize_name_tokens(name):
    if not isinstance(name, str):
        return []
    normalized = unicodedata.normalize("NFKD", name)
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.upper()
    normalized = re.sub(r"[^A-Z\s]", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized.split()


def name_key(name):
    tokens = normalize_name_tokens(name)
    if len(tokens) >= 2:
        return tokens[0], tokens[1]
    if len(tokens) == 1:
        return tokens[0], ""
    return "", ""


def map_turno_location(property_group, property_alias):
    combined = f"{property_group or ''} {property_alias or ''}".upper()
    if "MANGO" in combined:
        return "Mango Villas"
    if "DAMISELA" in combined:
        return "Casa Damisela"
    return "Other"


def process_timesheet(csv_file, output_excel, turno_csv, rates_csv=None):
    """Process timesheet and turno CSVs into an Excel workbook.

    At least one of csv_file or turno_csv must be provided.
    Returns a tuple (message, warnings) on success.
    Raises FileNotFoundError or ValueError on fatal errors.
    """
    warnings = []

    has_timeclock = bool(csv_file and csv_file.strip())
    has_turno = bool(turno_csv and turno_csv.strip())

    if not has_timeclock and not has_turno:
        raise ValueError("At least one input file (timeclock CSV or turno CSV) is required.")

    if has_timeclock and not os.path.exists(csv_file):
        raise FileNotFoundError(f"Input file not found: {csv_file}")
    if has_turno and not os.path.exists(turno_csv):
        raise FileNotFoundError(f"Turno file not found: {turno_csv}")

    # Reference file for rates lookup and date parsing
    date_source_file = csv_file if has_timeclock else turno_csv

    # Load rates CSV (ID,RATE,START,EXTRA,DETAILS)
    # Use explicit path if provided, otherwise look in parent folder of input CSV
    if rates_csv and os.path.exists(rates_csv):
        rates_file = rates_csv
    else:
        csv_dir = os.path.dirname(date_source_file)
        parent_dir = os.path.abspath(os.path.join(csv_dir, os.pardir))
        rates_file = os.path.join(parent_dir, "timesheet-rates.csv")
    if os.path.exists(rates_file):
        rates_df = pd.read_csv(rates_file)
        rates_df['START'] = pd.to_datetime(rates_df['START'], errors='coerce')
        rates_dict = {}
        rates_by_name = {}  # name_key -> ID for turno-only matching
        for _, row in rates_df.iterrows():
            rates_dict[row['ID']] = {
                'RATE': row['RATE'],
                'START': row['START'],
                'EXTRA': row['EXTRA'],
                'DETAILS': row['DETAILS'] if 'DETAILS' in row else ''
            }
            # Build name-based lookup if NAME column exists and is filled
            if 'NAME' in rates_df.columns and pd.notna(row.get('NAME')) and str(row['NAME']).strip():
                nk = name_key(str(row['NAME']).strip())
                if nk[0]:
                    rates_by_name[nk] = row['ID']
    else:
        warnings.append("'timesheet-rates.csv' not found. All rates will default to $0.")
        rates_dict = {}
        rates_by_name = {}

    # Process timeclock CSV
    persons = {}
    if has_timeclock:
        # Load main timesheet CSV
        df = pd.read_csv(csv_file)
        df.columns = [col.strip() for col in df.columns]

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
            warnings.append("Some datetime entries could not be parsed and will be skipped.")
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

    # Load and map turno CSV rows to people/locations
    turno_events = {}
    if has_turno:
        turno_df = pd.read_csv(turno_csv)
        turno_df.columns = [col.strip() for col in turno_df.columns]

        # Normalize column names to expected casing (CSV exports vary)
        expected_cols = {
            "Teammate": "Teammate",
            "Start Date & Time": "Start Date & Time",
            "End Date & Time": "End Date & Time",
            "Cleaning Price": "Cleaning Price",
            "Property Alias": "Property Alias",
            "Property Group": "Property Group",
        }
        col_lower_map = {col.lower(): col for col in turno_df.columns}
        rename_map = {}
        for expected in expected_cols:
            if expected not in turno_df.columns and expected.lower() in col_lower_map:
                rename_map[col_lower_map[expected.lower()]] = expected
        if rename_map:
            turno_df.rename(columns=rename_map, inplace=True)

        required_turno_cols = list(expected_cols.keys())
        missing_turno_cols = [col for col in required_turno_cols if col not in turno_df.columns]
        if missing_turno_cols:
            raise ValueError(f"Turno file missing columns: {', '.join(missing_turno_cols)}")

        turno_df["Start Dt"] = pd.to_datetime(turno_df["Start Date & Time"], errors="coerce")
        turno_df["End Dt"] = pd.to_datetime(turno_df["End Date & Time"], errors="coerce")
        turno_df["Cleaning Price"] = pd.to_numeric(turno_df["Cleaning Price"], errors="coerce").fillna(0)

        # Split pay when multiple people are assigned to the same property on the same date
        turno_df["Job Date"] = turno_df["Start Dt"].dt.date
        for (_, _), group in turno_df.groupby(["Property Alias", "Job Date"]):
            if len(group) > 1:
                # Use max price as the total project price (each row carries the full price)
                total_price = group["Cleaning Price"].max()
                split_price = round(total_price / len(group), 2)
                turno_df.loc[group.index, "Cleaning Price"] = split_price

        person_key_map = {}
        for person_key in persons.keys():
            first, last1 = name_key(person_key[1])
            if not first or not last1:
                continue
            person_key_map.setdefault((first, last1), []).append(person_key)

        for idx, row in turno_df.iterrows():
            teammate_val = row.get("Teammate", "")
            if pd.isna(teammate_val):
                warnings.append(f"Unmatched turno row {idx + 2}: missing teammate name")
                continue
            teammate = str(teammate_val).strip()
            first, last1 = name_key(teammate)
            if not first or not last1:
                warnings.append(f"Unmatched turno row {idx + 2}: invalid teammate name '{teammate}'")
                continue

            candidates = person_key_map.get((first, last1), [])
            if len(candidates) == 0:
                if has_timeclock:
                    warnings.append(f"Unmatched turno row {idx + 2}: teammate '{teammate}' not found in timesheet")
                    continue
                else:
                    # Turno-only mode: create person entry from teammate name
                    # Try to resolve Person ID from rates CSV by name
                    resolved_id = rates_by_name.get((first, last1), teammate)
                    person_key = (resolved_id, teammate)
                    if person_key not in persons:
                        persons[person_key] = pd.DataFrame(
                            columns=['Location', 'Date', 'Start', 'End', 'Hours', 'Details']
                        )
                        person_key_map.setdefault((first, last1), []).append(person_key)
                    candidates = person_key_map[(first, last1)]
            if len(candidates) > 1:
                candidate_labels = ", ".join([f"{pid} {pname}" for pid, pname in candidates])
                warnings.append(
                    f"Unmatched turno row {idx + 2}: teammate '{teammate}' matches multiple people ({candidate_labels})"
                )
                continue

            start_dt = row["Start Dt"]
            end_dt = row["End Dt"]
            rate_val = row["Cleaning Price"]
            if pd.isna(start_dt) or pd.isna(end_dt):
                warnings.append(f"Unmatched turno row {idx + 2}: missing start/end time for '{teammate}'")
                continue
            if pd.isna(rate_val):
                rate_val = 0

            property_alias_val = row.get("Property Alias", "")
            property_group_val = row.get("Property Group", "")
            property_alias = "" if pd.isna(property_alias_val) else str(property_alias_val).strip()
            property_group = "" if pd.isna(property_group_val) else str(property_group_val).strip()
            location = map_turno_location(property_group, property_alias)

            person_key = candidates[0]
            hours_worked = round((end_dt - start_dt).total_seconds() / 3600, 2)
            if hours_worked < 0.25 or hours_worked > 5:
                hours_worked = 2.0
            event = {
                "date": start_dt.strftime("%m/%d/%Y"),
                "start": start_dt.strftime("%H:%M:%S"),
                "end": end_dt.strftime("%H:%M:%S"),
                "hours": hours_worked,
                "rate": float(rate_val),
                "details": property_alias,
                "label": property_alias or "Details here",
                "start_dt": start_dt,
            }
            turno_events.setdefault(
                person_key, {"Mango Villas": [], "Casa Damisela": [], "Other": []}
            )
            turno_events[person_key][location].append(event)

        for location_groups in turno_events.values():
            for events in location_groups.values():
                events.sort(key=lambda item: item["start_dt"])

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
            if str(person_id) == str(person_name):
                sheet_name = str(person_name)
            else:
                sheet_name = f"{person_id} - {person_name}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            # Header and formatting
            # Write person info row with light green background
            if str(person_id) == str(person_name):
                worksheet.write(0, 0, f"Name: {person_name}", light_green_text_format)
            else:
                worksheet.write(0, 0, f"Person ID: {person_id}, Name: {person_name}", light_green_text_format)
            worksheet.write(0, 1, '', light_green_text_format)
            # Insert row for timesheet period (parsed from filename)
            match = re.search(r'(\d{2}-\d{2}-\d{4})', os.path.basename(date_source_file))
            if match:
                end_date = datetime.strptime(match.group(1), '%m-%d-%Y')
                start_date = end_date - timedelta(days=6)
                worksheet.write(1, 0, f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                worksheet.write_blank(2, 0, None)
            worksheet.set_column('A:D', 20)
            worksheet.set_column('E:E', 10)  # Hours
            worksheet.set_column('F:F', 12)  # Rate $
            worksheet.set_column('G:G', 25)  # Details

            def write_location_section(start_row, header_title, placeholders=None, data_rows=None):
                placeholders = placeholders or ["Apt X", "Apt X"]
                data_rows = data_rows or []
                worksheet.write(start_row, 0, header_title, header_format)
                section_headers = ['Date', 'Start', 'End', 'Hours', 'Rate $', 'Details']
                for col, name in enumerate(section_headers, start=1):
                    worksheet.write(start_row, col, name, header_format)

                data_start = start_row + 1
                row_count = max(len(placeholders), len(data_rows))
                if row_count == 0:
                    row_count = 1

                for offset in range(row_count):
                    label = placeholders[offset] if offset < len(placeholders) else ""
                    if offset < len(data_rows):
                        row_data = data_rows[offset]
                        label = row_data.get("label", label)
                        worksheet.write(data_start + offset, 1, row_data.get("date", ""))
                        worksheet.write(data_start + offset, 2, row_data.get("start", ""))
                        worksheet.write(data_start + offset, 3, row_data.get("end", ""))
                        worksheet.write_number(data_start + offset, 4, row_data.get("hours", 0))
                        worksheet.write_number(data_start + offset, 5, row_data.get("rate", 0), currency_format)
                    worksheet.write(data_start + offset, 0, label)

                total_row = data_start + row_count
                first_excel_row = data_start + 1
                last_excel_row = data_start + row_count
                worksheet.write(total_row, 0, 'Total $')
                worksheet.write_formula(total_row, 4, f"=SUM(E{first_excel_row}:E{last_excel_row})")
                worksheet.write_formula(total_row, 5, f"=SUM(F{first_excel_row}:F{last_excel_row})", currency_format)
                return {
                    "total_row": total_row,
                    "next_start": total_row + 2,
                    "data_start_excel": first_excel_row,
                    "data_end_excel": last_excel_row,
                }

            person_turno = turno_events.get((person_id, person_name), {})

            current_section_row = 3
            section_totals_excel_rows = []
            section_hours_excel_rows = []
            section_definitions = [
                ("Mango Villas", ["Apt X", "Apt X"]),
                ("Casa Damisela", ["Apt X", "Apt X"]),
                ("Other", ["Details here", ""])
            ]
            section_clean_ranges = []
            for section_header, placeholders in section_definitions:
                data_rows = person_turno.get(section_header, [])
                section_info = write_location_section(
                    current_section_row, section_header, placeholders, data_rows
                )
                section_totals_excel_rows.append(section_info["total_row"] + 1)
                section_hours_excel_rows.append(section_info["total_row"] + 1)
                section_clean_ranges.append(
                    (section_info["data_start_excel"], section_info["data_end_excel"])
                )
                current_section_row = section_info["next_start"]

            section_hours_refs = [f"E{r}" for r in section_hours_excel_rows]

            # Summary header directly after the location sections
            summary_header_row = current_section_row
            worksheet.write(summary_header_row, 0, "Summary", summary_header_format)
            worksheet.merge_range(summary_header_row, 0, summary_header_row, 5, "Summary", summary_header_format)

            total_hours_block_row = summary_header_row + 1
            worksheet.write(total_hours_block_row, 0, "Total Hours")
            hours_formula = "=" + "+".join(section_hours_refs) if section_hours_refs else "=0"
            worksheet.write_formula(total_hours_block_row, 4, hours_formula)

            extras_row_idx = summary_header_row + 2
            subtotal_row_idx = extras_row_idx + 1
            worksheet.write(subtotal_row_idx, 0, "Subtotal $")

            extras_excel_row = extras_row_idx + 1

            # Rates lookup (for extras/withholding)
            start_date = None
            extra_val = 0
            details_val = ''
            if person_id in rates_dict:
                start_date = rates_dict[person_id]['START']
                extra_val = rates_dict[person_id]['EXTRA']
                details_val = str(rates_dict[person_id]['DETAILS']) if pd.notna(rates_dict[person_id]['DETAILS']) else ""

            # Extras and Annual withheld exclusion
            worksheet.write(extras_row_idx, 0, "Extras $")

            today = datetime.now()
            show_red = False
            if start_date is not None:
                if (today - start_date).days < 28 or today.month == 1:
                    # Within 4 weeks or January
                    extras_amount = extra_val
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
            # Reviewed marker cell to the right of Total $
            review_cell = f"F{total_dollar_idx + 1}"
            worksheet.write_blank(total_dollar_idx, 5, None)
            worksheet.data_validation(total_dollar_idx, 5, total_dollar_idx, 5, {
                "validate": "list",
                "source": ["", "y"],
                "input_title": "Reviewed?",
                "input_message": "Choose 'y' once this sheet is reviewed.",
            })

            # Subtotal: section totals + extras
            section_total_refs = [f"F{row}" for row in section_totals_excel_rows]
            subtotal_parts = [*section_total_refs, f"E{extras_excel_row}"]
            subtotal_formula = "=" + " + ".join(subtotal_parts)
            worksheet.write_formula(subtotal_row_idx, 4, subtotal_formula, currency_format)

            # Capture summary references
            hours_cell = f"={quote_sheetname(sheet_name)}!E{total_hours_block_row + 1}"
            total_cell = f"={quote_sheetname(sheet_name)}!E{total_dollar_idx + 1}"
            withheld_cell = f"={quote_sheetname(sheet_name)}!E{withheld_row_idx + 1}"
            reviewed_cell = f"=IF({quote_sheetname(sheet_name)}!{review_cell}=\"y\",\"y\",\"\")"

            sheet_ref = quote_sheetname(sheet_name)
            section_clean_counts = [
                f'COUNTIF({sheet_ref}!F{start}:F{end},"<>")' for start, end in section_clean_ranges
            ]
            cleans_formula = "=" + "+".join(section_clean_counts) if section_clean_counts else "=0"

            summary_entries.append((sheet_name, hours_cell, cleans_formula, total_cell, withheld_cell, reviewed_cell))

        # Populate summary sheet
        summary_sheet.write(0, 0, "Person", summary_header_format)
        summary_sheet.write(0, 1, "Total Hours", summary_header_format)
        summary_sheet.write(0, 2, "Total Cleans", summary_header_format)
        summary_sheet.write(0, 3, "Total $", summary_header_format)
        summary_sheet.write(0, 4, "Withheld $", summary_header_format)
        summary_sheet.write(0, 5, "Pay/Hour", summary_header_format)
        summary_sheet.write(0, 6, "Pay/Job", summary_header_format)
        summary_sheet.write(0, 7, "Reviewed", summary_header_format)
        for idx, (label, hours_ref, cleans_ref, total_ref, withheld_ref, reviewed_ref) in enumerate(summary_entries, start=1):
            excel_row = idx + 1
            summary_sheet.write(idx, 0, label)
            summary_sheet.write_formula(idx, 1, hours_ref)
            summary_sheet.write_formula(idx, 2, cleans_ref)
            summary_sheet.write_formula(idx, 3, total_ref, currency_format)
            summary_sheet.write_formula(idx, 4, withheld_ref, currency_format)
            summary_sheet.write_formula(idx, 5, f'=IFERROR(D{excel_row}/B{excel_row},"")', currency_format)
            summary_sheet.write_formula(idx, 6, f'=IFERROR(D{excel_row}/C{excel_row},"")', currency_format)
            summary_sheet.write_formula(idx, 7, reviewed_ref)
        if summary_entries:
            total_row = len(summary_entries) + 1
            summary_sheet.write(total_row, 0, "All sheets total", summary_header_format)
            summary_sheet.write_formula(total_row, 1, f"=SUM(B2:B{total_row})")
            summary_sheet.write_formula(total_row, 2, f"=SUM(C2:C{total_row})")
            summary_sheet.write_formula(total_row, 3, f"=SUM(D2:D{total_row})", light_green_currency_format)
            summary_sheet.write_formula(total_row, 4, f"=SUM(E2:E{total_row})", currency_format)
        # Set each column to the minimum width needed to display its content without wrapping
        all_labels = ([label for label, *_ in summary_entries] + ["All sheets total"]) if summary_entries else ["All sheets total"]
        col_a_width = max(len(s) for s in all_labels) + 2
        summary_sheet.set_column(0, 0, col_a_width)
        summary_sheet.set_column(1, 1, 13)   # Total Hours
        summary_sheet.set_column(2, 2, 14)   # Total Cleans
        summary_sheet.set_column(3, 3, 12)   # Total $
        summary_sheet.set_column(4, 4, 12)   # Withheld $
        summary_sheet.set_column(5, 5, 12)   # Pay/Hour
        summary_sheet.set_column(6, 6, 11)   # Pay/Job
        summary_sheet.set_column(7, 7, 10)   # Reviewed

    message = f"Excel file '{output_excel}' created successfully."
    return message, warnings

if __name__ == "__main__":
    if len(sys.argv) < 3 or len(sys.argv) > 4:
        print("Usage: python export-timesheet.py <output_excel> <input_csv> [turno_csv]")
        sys.exit(1)

    output_excel = sys.argv[1]
    input_csv = sys.argv[2]
    turno_csv = sys.argv[3] if len(sys.argv) > 3 else None
    try:
        message, warn_list = process_timesheet(input_csv, output_excel, turno_csv)
        for w in warn_list:
            print(f"Warning: {w}")
        print(message)
    except (FileNotFoundError, ValueError) as e:
        print(e)
        sys.exit(1)
