import argparse
import os
import re
import sys
import unicodedata
import warnings as py_warnings
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
from xlsxwriter.utility import quote_sheetname


LOCAL_TZ = ZoneInfo("America/Puerto_Rico")
LOCATION_BUCKETS = ["Mango Villas", "Casa Damisela", "MARU", "Other"]
AMBIGUOUS_PERSON = object()


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


def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def id_key(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def map_turno_location(property_group, property_alias):
    combined = f"{property_group or ''} {property_alias or ''}".upper()
    if "MANGO" in combined:
        return "Mango Villas"
    if "DAMISELA" in combined:
        return "Casa Damisela"
    if "MARU" in combined:
        return "MARU"
    # MARU rooms are exported with a blank Property Group and aliases like "ROOM ONE".
    if re.search(r"\bROOM\s+(ONE|TWO|THREE|FOUR|FIVE|[1-5])\b", combined):
        return "MARU"
    return "Other"


def _normalize_expected_columns(df, expected_names):
    """Rename columns case-insensitively to the expected display names."""
    col_lower_map = {col.lower().strip(): col for col in df.columns}
    rename_map = {}
    for expected in expected_names:
        lower = expected.lower()
        if expected not in df.columns and lower in col_lower_map:
            rename_map[col_lower_map[lower]] = expected
    if rename_map:
        df.rename(columns=rename_map, inplace=True)


def _find_date_in_paths(output_excel, input_paths):
    for path in [output_excel, *input_paths]:
        if not path:
            continue
        match = re.search(r"(\d{2}-\d{2}-\d{4})", os.path.basename(path))
        if match:
            return datetime.strptime(match.group(1), "%m-%d-%Y")
    return None


def _find_default_rates_file(source_file):
    if not source_file:
        return None
    candidate = os.path.abspath(os.path.dirname(source_file))
    for _ in range(6):
        rates_file = os.path.join(candidate, "timesheet-rates.csv")
        if os.path.exists(rates_file):
            return rates_file
        parent = os.path.dirname(candidate)
        if parent == candidate:
            break
        candidate = parent
    return None


def _parse_datetime_series(series, formats):
    text = series.astype(str).str.strip()
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    for fmt in formats:
        mask = parsed.isna()
        if not mask.any():
            break
        parsed.loc[mask] = pd.to_datetime(text[mask], format=fmt, errors="coerce")

    mask = parsed.isna() & text.ne("") & text.str.lower().ne("nan")
    if mask.any():
        with py_warnings.catch_warnings():
            py_warnings.simplefilter("ignore", UserWarning)
            parsed.loc[mask] = pd.to_datetime(text[mask], errors="coerce")
    return parsed


def _parse_money_value(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    is_parenthesized_negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()").replace("$", "").replace(",", "").strip()
    parsed = pd.to_numeric(pd.Series([text]), errors="coerce").iloc[0]
    if pd.isna(parsed):
        return None
    amount = float(parsed)
    if is_parenthesized_negative:
        return -amount
    return amount


def _normalize_reimbursable(value):
    text = clean_value(value).lower()
    if text in {"yes", "y", "true", "1"}:
        return "Yes", True, False
    if text in {"", "no", "n", "false", "0"}:
        return "No", False, False
    return "No", False, True


def _load_rates(rates_csv, source_file, warnings):
    if rates_csv and os.path.exists(rates_csv):
        rates_file = rates_csv
    else:
        rates_file = _find_default_rates_file(source_file)
        if rates_csv and not os.path.exists(rates_csv):
            warnings.append(f"Selected rates file not found: {rates_csv}")

    if not rates_file or not os.path.exists(rates_file):
        warnings.append("'timesheet-rates.csv' not found. Hourly rates and extras default to $0.")
        return {}, {}

    rates_df = pd.read_csv(rates_file)
    rates_df.columns = [col.strip() for col in rates_df.columns]

    required_cols = ["ID", "RATE", "START", "EXTRA"]
    missing_cols = [col for col in required_cols if col not in rates_df.columns]
    if missing_cols:
        warnings.append(
            f"Rates file missing columns ({', '.join(missing_cols)}). Hourly rates and extras default to $0."
        )
        return {}, {}

    rates_df["START"] = pd.to_datetime(rates_df["START"], errors="coerce")
    rates_df["RATE"] = pd.to_numeric(rates_df["RATE"], errors="coerce").fillna(0)
    rates_df["EXTRA"] = pd.to_numeric(rates_df["EXTRA"], errors="coerce").fillna(0)

    rates_dict = {}
    rates_by_name = {}
    for _, row in rates_df.iterrows():
        person_id = id_key(row["ID"])
        if not person_id:
            continue
        person_name = clean_value(row.get("NAME", ""))
        rates_dict[person_id] = {
            "NAME": person_name,
            "RATE": float(row["RATE"]),
            "START": row["START"],
            "EXTRA": float(row["EXTRA"]),
            "DETAILS": clean_value(row.get("DETAILS", "")),
        }
        if person_name:
            nk = name_key(person_name)
            if nk[0]:
                rates_by_name.setdefault(nk, [])
                if person_id not in rates_by_name[nk]:
                    rates_by_name[nk].append(person_id)

    return rates_dict, rates_by_name


def _resolve_rate_id_by_name(person_name, rates_by_name, warnings, context):
    nk = name_key(person_name)
    if not nk[0]:
        return None
    matches = rates_by_name.get(nk, [])
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        warnings.append(
            f"{context}: name '{person_name}' matches multiple rate rows ({', '.join(matches)}); using $0 rate."
        )
    return None


def _get_rate_info(person_id, person_name, rates_dict, rates_by_name):
    normalized_id = id_key(person_id)
    if normalized_id in rates_dict:
        return rates_dict[normalized_id], True

    resolved_id = _resolve_rate_id_by_name(person_name, rates_by_name, [], "rate lookup")
    if resolved_id and resolved_id in rates_dict:
        return rates_dict[resolved_id], True

    return {
        "NAME": "",
        "RATE": 0.0,
        "START": pd.NaT,
        "EXTRA": 0.0,
        "DETAILS": "",
    }, False


def _ensure_person(persons, hourly_events, turno_events, person_key, expense_events=None):
    persons.setdefault(person_key, True)
    hourly_events.setdefault(person_key, [])
    turno_events.setdefault(person_key, {bucket: [] for bucket in LOCATION_BUCKETS})
    if expense_events is not None:
        expense_events.setdefault(person_key, [])


def _find_existing_person_by_name(persons, person_name, warnings, context):
    nk = name_key(person_name)
    if not nk[0]:
        return None
    matches = [person_key for person_key in persons if name_key(person_key[1]) == nk]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        labels = ", ".join([f"{pid} {pname}" for pid, pname in matches])
        warnings.append(f"{context}: name '{person_name}' matches multiple people ({labels}); row skipped.")
        return AMBIGUOUS_PERSON
    return None


def _person_from_name(person_name, persons, rates_by_name, warnings, context):
    existing = _find_existing_person_by_name(persons, person_name, warnings, context)
    if existing is AMBIGUOUS_PERSON:
        return None
    if existing:
        return existing

    resolved_id = _resolve_rate_id_by_name(person_name, rates_by_name, warnings, context)
    if resolved_id:
        return resolved_id, person_name

    return person_name, person_name


def _parse_timeclock(csv_file, persons, hourly_events, turno_events, warnings):
    df = pd.read_csv(csv_file)
    df.columns = [col.strip() for col in df.columns]
    required_cols = ["Person ID", "Person Name", "Punch Date", "Attendance record"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Timeclock file missing columns: {', '.join(missing_cols)}")

    if df.empty:
        warnings.append(f"Timeclock file '{os.path.basename(csv_file)}' has no rows.")
        return

    datetime_str = df["Punch Date"].astype(str).str.strip() + " " + df["Attendance record"].astype(str).str.strip()
    df["Datetime"] = pd.to_datetime(datetime_str, format="%Y-%m-%d %H:%M:%S", errors="coerce")
    mask_failed = df["Datetime"].isna()
    if mask_failed.any():
        df.loc[mask_failed, "Datetime"] = pd.to_datetime(
            datetime_str[mask_failed], format="%m/%d/%Y %H:%M:%S", errors="coerce"
        )

    if df["Datetime"].isna().any():
        skipped = int(df["Datetime"].isna().sum())
        warnings.append(f"{skipped} timeclock rows had unparseable dates and were skipped.")
        df = df.dropna(subset=["Datetime"])

    if df.empty:
        warnings.append(f"Timeclock file '{os.path.basename(csv_file)}' had no usable rows.")
        return

    for (person_id, person_name, punch_date), group in df.groupby(["Person ID", "Person Name", "Punch Date"]):
        group_sorted = group.sort_values(by="Datetime")
        check_in = group_sorted.iloc[0]["Datetime"]
        check_out = group_sorted.iloc[-1]["Datetime"]
        hours_worked = round((check_out - check_in).total_seconds() / 3600, 2)
        if hours_worked <= 0:
            warnings.append(f"Timeclock row for '{person_name}' on {punch_date} has no positive hours; row skipped.")
            continue

        person_key = (id_key(person_id), clean_value(person_name))
        _ensure_person(persons, hourly_events, turno_events, person_key)
        hourly_events[person_key].append({
            "source": "Timeclock",
            "location": "Maru",
            "date": check_in.strftime("%m/%d/%Y"),
            "start": check_in.strftime("%H:%M:%S"),
            "end": check_out.strftime("%H:%M:%S"),
            "hours": hours_worked,
            "details": "",
            "start_dt": check_in,
        })


def _parse_notion(notion_csv, persons, hourly_events, turno_events, rates_by_name, warnings):
    notion_df = pd.read_csv(notion_csv)
    notion_df.columns = [col.strip() for col in notion_df.columns]
    expected_cols = [
        "Date",
        "Status",
        "Category",
        "Team Member",
        "Person",
        "Property",
        "Start Time (UTC)",
        "End Time (UTC)",
        "Hours (calc)",
        "Notes",
        "Time Log URL",
    ]
    _normalize_expected_columns(notion_df, expected_cols)

    required_cols = ["Start Time (UTC)", "End Time (UTC)", "Hours (calc)"]
    missing_cols = [col for col in required_cols if col not in notion_df.columns]
    if missing_cols:
        raise ValueError(f"Notion file missing columns: {', '.join(missing_cols)}")

    if "Person" not in notion_df.columns and "Team Member" not in notion_df.columns:
        raise ValueError("Notion file missing columns: Person or Team Member")

    if notion_df.empty:
        warnings.append(f"Notion file '{os.path.basename(notion_csv)}' has no rows.")
        return

    fallback_rows = 0
    for idx, row in notion_df.iterrows():
        row_number = idx + 2
        person_name = clean_value(row.get("Person", ""))
        if not person_name:
            person_name = clean_value(row.get("Team Member", ""))
            if person_name:
                fallback_rows += 1

        if not person_name:
            warnings.append(f"Notion row {row_number}: missing Person and Team Member; row skipped.")
            continue

        start_utc = pd.to_datetime(row.get("Start Time (UTC)", ""), utc=True, errors="coerce")
        end_utc = pd.to_datetime(row.get("End Time (UTC)", ""), utc=True, errors="coerce")
        if pd.isna(start_utc) or pd.isna(end_utc):
            warnings.append(f"Notion row {row_number}: missing or invalid start/end time for '{person_name}'; row skipped.")
            continue

        hours_value = pd.to_numeric(pd.Series([row.get("Hours (calc)", "")]), errors="coerce").iloc[0]
        if pd.isna(hours_value):
            hours_value = round((end_utc - start_utc).total_seconds() / 3600, 2)
        hours_value = round(float(hours_value), 2)
        if hours_value <= 0:
            warnings.append(f"Notion row {row_number}: hours must be positive for '{person_name}'; row skipped.")
            continue

        person_key = _person_from_name(person_name, persons, rates_by_name, warnings, f"Notion row {row_number}")
        if person_key is None:
            continue
        _ensure_person(persons, hourly_events, turno_events, person_key)

        start_local = start_utc.tz_convert(LOCAL_TZ)
        end_local = end_utc.tz_convert(LOCAL_TZ)

        property_name = clean_value(row.get("Property", ""))
        category = clean_value(row.get("Category", ""))
        status = clean_value(row.get("Status", ""))
        notes = clean_value(row.get("Notes", ""))
        url = clean_value(row.get("Time Log URL", ""))
        details = "; ".join(
            part for part in [
                f"Status: {status}" if status else "",
                f"Category: {category}" if category else "",
                notes,
                url,
            ] if part
        )

        hourly_events[person_key].append({
            "source": "Notion",
            "location": property_name or category or "Notion",
            "date": start_local.strftime("%m/%d/%Y"),
            "start": start_local.strftime("%H:%M:%S"),
            "end": end_local.strftime("%H:%M:%S"),
            "hours": hours_value,
            "details": details,
            "start_dt": start_local.to_pydatetime(),
        })

    if fallback_rows:
        warnings.append(f"{fallback_rows} Notion rows used Team Member because Person was blank.")

    for rows in hourly_events.values():
        rows.sort(key=lambda item: item["start_dt"])


def _parse_turno(turno_csv, persons, hourly_events, turno_events, rates_by_name, warnings):
    turno_df = pd.read_csv(turno_csv)
    turno_df.columns = [col.strip() for col in turno_df.columns]

    expected_cols = [
        "Teammate",
        "Start Date & Time",
        "End Date & Time",
        "Cleaning Price",
        "Property Alias",
        "Property Group",
    ]
    _normalize_expected_columns(turno_df, expected_cols)

    missing_cols = [col for col in expected_cols if col not in turno_df.columns]
    if missing_cols:
        raise ValueError(f"Turno file missing columns: {', '.join(missing_cols)}")

    if turno_df.empty:
        warnings.append(f"Turno file '{os.path.basename(turno_csv)}' has no rows.")
        return

    turno_formats = ["%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d %H:%M:%S"]
    turno_df["Start Dt"] = _parse_datetime_series(turno_df["Start Date & Time"], turno_formats)
    turno_df["End Dt"] = _parse_datetime_series(turno_df["End Date & Time"], turno_formats)
    turno_df["Cleaning Price"] = pd.to_numeric(turno_df["Cleaning Price"], errors="coerce").fillna(0)
    turno_df["Job Date"] = turno_df["Start Dt"].dt.date

    # Split pay when multiple people are assigned to the same property on the same date.
    for (_, _), group in turno_df.groupby(["Property Alias", "Job Date"]):
        if len(group) > 1:
            total_price = group["Cleaning Price"].max()
            split_price = round(total_price / len(group), 2)
            turno_df.loc[group.index, "Cleaning Price"] = split_price

    for idx, row in turno_df.iterrows():
        row_number = idx + 2
        teammate = clean_value(row.get("Teammate", ""))
        if not teammate:
            warnings.append(f"Turno row {row_number}: missing teammate name; row skipped.")
            continue
        first, last1 = name_key(teammate)
        if not first or not last1:
            warnings.append(f"Turno row {row_number}: invalid teammate name '{teammate}'; row skipped.")
            continue

        person_key = _person_from_name(teammate, persons, rates_by_name, warnings, f"Turno row {row_number}")
        if person_key is None:
            continue
        _ensure_person(persons, hourly_events, turno_events, person_key)

        start_dt = row["Start Dt"]
        end_dt = row["End Dt"]
        if pd.isna(start_dt) or pd.isna(end_dt):
            warnings.append(f"Turno row {row_number}: missing start/end time for '{teammate}'; row skipped.")
            continue

        property_alias = clean_value(row.get("Property Alias", ""))
        property_group = clean_value(row.get("Property Group", ""))
        location = map_turno_location(property_group, property_alias)

        hours_worked = round((end_dt - start_dt).total_seconds() / 3600, 2)
        if hours_worked < 0.25 or hours_worked > 5:
            hours_worked = 2.0

        event = {
            "date": start_dt.strftime("%m/%d/%Y"),
            "start": start_dt.strftime("%H:%M:%S"),
            "end": end_dt.strftime("%H:%M:%S"),
            "hours": hours_worked,
            "rate": float(row["Cleaning Price"]),
            "details": property_group,
            "label": property_alias,
            "start_dt": start_dt,
        }
        turno_events[person_key][location].append(event)

    for location_groups in turno_events.values():
        for events in location_groups.values():
            events.sort(key=lambda item: item["start_dt"])


def _parse_expenses(expenses_csv, persons, hourly_events, turno_events, expense_events, rates_by_name, warnings):
    expenses_df = pd.read_csv(expenses_csv)
    expenses_df.columns = [col.strip() for col in expenses_df.columns]

    expected_cols = [
        "Expensed By",
        "Date",
        "Category",
        "Expense",
        "Vendor",
        "Property",
        "Unit",
        "Amount",
        "Payment Method",
        "Reimbursable",
        "Approved By",
        "Notes",
        "Expense URL",
    ]
    _normalize_expected_columns(expenses_df, expected_cols)

    required_cols = ["Expensed By", "Date", "Expense", "Amount", "Reimbursable"]
    missing_cols = [col for col in required_cols if col not in expenses_df.columns]
    if missing_cols:
        raise ValueError(f"Expenses file missing columns: {', '.join(missing_cols)}")

    if expenses_df.empty:
        warnings.append(f"Expenses file '{os.path.basename(expenses_csv)}' has no rows.")
        return

    expense_formats = ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"]
    expenses_df["Expense Dt"] = _parse_datetime_series(expenses_df["Date"], expense_formats)

    for idx, row in expenses_df.iterrows():
        row_number = idx + 2
        expensed_by = clean_value(row.get("Expensed By", ""))
        if not expensed_by:
            warnings.append(f"Expenses row {row_number}: missing Expensed By; row skipped.")
            continue

        expense_dt = row["Expense Dt"]
        if pd.isna(expense_dt):
            warnings.append(f"Expenses row {row_number}: missing or invalid date for '{expensed_by}'; row skipped.")
            continue

        amount = _parse_money_value(row.get("Amount", ""))
        if amount is None:
            warnings.append(f"Expenses row {row_number}: missing or invalid amount for '{expensed_by}'; row skipped.")
            continue

        reimbursement_label, is_reimbursable, unknown_reimbursement = _normalize_reimbursable(
            row.get("Reimbursable", "")
        )
        if unknown_reimbursement:
            warnings.append(
                f"Expenses row {row_number}: Reimbursable value '{clean_value(row.get('Reimbursable', ''))}' "
                "was not recognized; treating it as No."
            )

        person_key = _person_from_name(expensed_by, persons, rates_by_name, warnings, f"Expenses row {row_number}")
        if person_key is None:
            continue
        _ensure_person(persons, hourly_events, turno_events, person_key, expense_events)

        property_name = clean_value(row.get("Property", ""))
        unit = clean_value(row.get("Unit", ""))
        property_label = property_name
        if unit:
            property_label = f"{property_name} - {unit}" if property_name else unit

        payment_method = clean_value(row.get("Payment Method", ""))
        approved_by = clean_value(row.get("Approved By", ""))
        notes = clean_value(row.get("Notes", ""))
        url = clean_value(row.get("Expense URL", ""))
        details = "; ".join(
            part for part in [
                f"Payment: {payment_method}" if payment_method else "",
                f"Approved by: {approved_by}" if approved_by else "",
                notes,
                url,
            ] if part
        )

        expense_events[person_key].append({
            "date": expense_dt.strftime("%m/%d/%Y"),
            "category": clean_value(row.get("Category", "")),
            "expense": clean_value(row.get("Expense", "")),
            "vendor": clean_value(row.get("Vendor", "")),
            "property": property_label,
            "amount": amount,
            "reimbursable": reimbursement_label,
            "is_reimbursable": is_reimbursable,
            "details": details,
            "expense_dt": expense_dt,
        })

    for rows in expense_events.values():
        rows.sort(key=lambda item: item["expense_dt"])


def _safe_sheet_name(base_name, used_names):
    safe_name = re.sub(r"[\[\]\:\*\?\/\\]", "-", base_name).strip() or "Employee"
    safe_name = safe_name[:31]
    candidate = safe_name
    suffix = 2
    while candidate in used_names:
        suffix_text = f" {suffix}"
        candidate = f"{safe_name[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def _parse_work_date(value):
    try:
        return datetime.strptime(str(value), "%m/%d/%Y").date()
    except (TypeError, ValueError):
        return None


def _count_worked_days(hourly_rows, turno_location_rows, period_start=None, period_end=None):
    period_start_date = period_start.date() if period_start else None
    period_end_date = period_end.date() if period_end else None
    worked_dates = set()

    def add_date(value):
        work_date = _parse_work_date(value)
        if not work_date:
            return
        if period_start_date and work_date < period_start_date:
            return
        if period_end_date and work_date > period_end_date:
            return
        worked_dates.add(work_date)

    for row in hourly_rows:
        add_date(row.get("date"))

    for rows in turno_location_rows.values():
        for row in rows:
            add_date(row.get("date"))

    return len(worked_dates)


def _has_turno_rows(turno_location_rows):
    return any(bool(rows) for rows in turno_location_rows.values())


def _has_hourly_source(hourly_rows, source):
    return any(row.get("source") == source for row in hourly_rows)


def _person_role(hourly_rows, turno_location_rows, expense_rows=None):
    has_turno_rows = _has_turno_rows(turno_location_rows)
    has_notion_rows = _has_hourly_source(hourly_rows, "Notion")
    has_timeclock_rows = _has_hourly_source(hourly_rows, "Timeclock")
    has_expense_rows = bool(expense_rows)

    if has_turno_rows and has_notion_rows:
        return "Housekeeping / Contractor"
    if has_turno_rows:
        return "Housekeeping"
    if has_notion_rows:
        return "Contractor"
    if has_timeclock_rows:
        return "Housekeeping"
    if has_expense_rows:
        return "Expenses"
    return ""


def _period_text(period_start, period_end):
    if period_start and period_end:
        return f"{period_start.strftime('%Y-%m-%d')} to {period_end.strftime('%Y-%m-%d')}"
    return ""


def _person_period(hourly_rows, turno_location_rows, period_end):
    if not period_end:
        return None, None, ""

    has_notion_rows = _has_hourly_source(hourly_rows, "Notion")
    has_turno_rows = _has_turno_rows(turno_location_rows)
    period_days = 14 if has_notion_rows and not has_turno_rows else 7
    if has_notion_rows and has_turno_rows:
        period_days = 14

    period_start = period_end - timedelta(days=period_days - 1)
    return period_start, period_end, _period_text(period_start, period_end)


def process_timesheet(csv_file, output_excel, turno_csv=None, rates_csv=None, notion_csv=None, expenses_csv=None):
    """Process payroll CSV inputs into an Excel workbook.

    csv_file is the legacy timeclock input. At least one of csv_file, turno_csv,
    notion_csv, or expenses_csv must be provided. Returns (message, warnings)
    on success.
    """
    warnings = []

    has_timeclock = bool(csv_file and str(csv_file).strip())
    has_turno = bool(turno_csv and str(turno_csv).strip())
    has_notion = bool(notion_csv and str(notion_csv).strip())
    has_expenses = bool(expenses_csv and str(expenses_csv).strip())

    if not has_timeclock and not has_turno and not has_notion and not has_expenses:
        raise ValueError("At least one input file (Notion, Turno, Timeclock, or Expenses CSV) is required.")

    if has_timeclock and not os.path.exists(csv_file):
        raise FileNotFoundError(f"Timeclock file not found: {csv_file}")
    if has_turno and not os.path.exists(turno_csv):
        raise FileNotFoundError(f"Turno file not found: {turno_csv}")
    if has_notion and not os.path.exists(notion_csv):
        raise FileNotFoundError(f"Notion file not found: {notion_csv}")
    if has_expenses and not os.path.exists(expenses_csv):
        raise FileNotFoundError(f"Expenses file not found: {expenses_csv}")

    input_files = [path for path in [notion_csv, csv_file, turno_csv, expenses_csv] if path]
    source_file = input_files[0]
    rates_dict, rates_by_name = _load_rates(rates_csv, source_file, warnings)

    persons = {}
    hourly_events = {}
    turno_events = {}
    expense_events = {}

    if has_timeclock:
        _parse_timeclock(csv_file, persons, hourly_events, turno_events, warnings)
    if has_notion:
        _parse_notion(notion_csv, persons, hourly_events, turno_events, rates_by_name, warnings)
    if has_turno:
        _parse_turno(turno_csv, persons, hourly_events, turno_events, rates_by_name, warnings)
    if has_expenses:
        _parse_expenses(expenses_csv, persons, hourly_events, turno_events, expense_events, rates_by_name, warnings)

    period_end = _find_date_in_paths(output_excel, input_files)
    missing_rate_people = set()

    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "border": 1})
        currency_format = workbook.add_format({"num_format": "$#,##0.00"})
        soft_red_format = workbook.add_format({"bg_color": "#FF9999", "num_format": "$#,##0.00"})
        light_green_text_format = workbook.add_format({"bg_color": "#CCFFCC"})
        light_green_currency_format = workbook.add_format({"bg_color": "#CCFFCC", "num_format": "$#,##0.00"})
        summary_header_format = workbook.add_format({"bold": True, "border": 1})

        summary_sheet = workbook.add_worksheet("Summary")
        writer.sheets["Summary"] = summary_sheet
        summary_entries = []
        used_sheet_names = {"Summary"}

        def write_hourly_section(worksheet, start_row, data_rows, hourly_rate):
            worksheet.write(start_row, 0, "Hourly Work", header_format)
            section_headers = ["Date", "Start", "End", "Hours", "Rate $", "Total $", "Details"]
            for col, name in enumerate(section_headers, start=1):
                worksheet.write(start_row, col, name, header_format)

            data_start = start_row + 1
            row_count = max(1, len(data_rows))
            for offset in range(row_count):
                excel_row = data_start + offset + 1
                if offset < len(data_rows):
                    row_data = data_rows[offset]
                    worksheet.write(data_start + offset, 0, row_data.get("location", ""))
                    worksheet.write(data_start + offset, 1, row_data.get("date", ""))
                    worksheet.write(data_start + offset, 2, row_data.get("start", ""))
                    worksheet.write(data_start + offset, 3, row_data.get("end", ""))
                    worksheet.write_number(data_start + offset, 4, row_data.get("hours", 0))
                    worksheet.write_number(data_start + offset, 5, hourly_rate, currency_format)
                    worksheet.write_formula(data_start + offset, 6, f"=E{excel_row}*F{excel_row}", currency_format)
                    worksheet.write(data_start + offset, 7, row_data.get("details", ""))

            total_row = data_start + row_count
            first_excel_row = data_start + 1
            last_excel_row = data_start + row_count
            worksheet.write(total_row, 0, "Total $")
            worksheet.write_formula(total_row, 4, f"=SUM(E{first_excel_row}:E{last_excel_row})")
            worksheet.write_formula(total_row, 6, f"=SUM(G{first_excel_row}:G{last_excel_row})", currency_format)
            return {
                "next_start": total_row + 2,
                "hours_total_cell": f"E{total_row + 1}",
                "dollar_total_cell": f"G{total_row + 1}",
            }

        def write_location_section(worksheet, start_row, header_title, data_rows=None):
            data_rows = data_rows or []
            worksheet.write(start_row, 0, header_title, header_format)
            section_headers = ["Date", "Start", "End", "Hours", "Rate $", "Details"]
            for col, name in enumerate(section_headers, start=1):
                worksheet.write(start_row, col, name, header_format)

            data_start = start_row + 1
            # Keep one blank row in empty sections for manual additions and valid SUM ranges.
            row_count = max(len(data_rows), 1)
            for offset in range(row_count):
                label = ""
                if offset < len(data_rows):
                    row_data = data_rows[offset]
                    label = row_data.get("label", "")
                    worksheet.write(data_start + offset, 1, row_data.get("date", ""))
                    worksheet.write(data_start + offset, 2, row_data.get("start", ""))
                    worksheet.write(data_start + offset, 3, row_data.get("end", ""))
                    worksheet.write_number(data_start + offset, 4, row_data.get("hours", 0))
                    worksheet.write_number(data_start + offset, 5, row_data.get("rate", 0), currency_format)
                    worksheet.write(data_start + offset, 6, row_data.get("details", ""))
                worksheet.write(data_start + offset, 0, label)

            total_row = data_start + row_count
            first_excel_row = data_start + 1
            last_excel_row = data_start + row_count
            worksheet.write(total_row, 0, "Total $")
            worksheet.write_formula(total_row, 4, f"=SUM(E{first_excel_row}:E{last_excel_row})")
            worksheet.write_formula(total_row, 5, f"=SUM(F{first_excel_row}:F{last_excel_row})", currency_format)
            return {
                "next_start": total_row + 2,
                "hours_total_cell": f"E{total_row + 1}",
                "dollar_total_cell": f"F{total_row + 1}",
                "clean_start_excel": first_excel_row,
                "clean_end_excel": last_excel_row,
            }

        def write_expense_section(worksheet, start_row, data_rows):
            worksheet.write(start_row, 0, "Expenses", header_format)
            section_headers = ["Date", "Category", "Expense", "Vendor", "Amount $", "Reimbursable", "Details"]
            for col, name in enumerate(section_headers, start=1):
                worksheet.write(start_row, col, name, header_format)

            data_start = start_row + 1
            row_count = max(1, len(data_rows))
            for offset in range(row_count):
                if offset < len(data_rows):
                    row_data = data_rows[offset]
                    worksheet.write(data_start + offset, 0, row_data.get("property", ""))
                    worksheet.write(data_start + offset, 1, row_data.get("date", ""))
                    worksheet.write(data_start + offset, 2, row_data.get("category", ""))
                    worksheet.write(data_start + offset, 3, row_data.get("expense", ""))
                    worksheet.write(data_start + offset, 4, row_data.get("vendor", ""))
                    worksheet.write_number(data_start + offset, 5, row_data.get("amount", 0), currency_format)
                    worksheet.write(data_start + offset, 6, row_data.get("reimbursable", "No"))
                    worksheet.write(data_start + offset, 7, row_data.get("details", ""))

            total_row = data_start + row_count
            first_excel_row = data_start + 1
            last_excel_row = data_start + row_count
            worksheet.write(total_row, 0, "Total reimbursable $")
            worksheet.write_formula(
                total_row,
                5,
                f'=SUMIF(G{first_excel_row}:G{last_excel_row},"Yes",F{first_excel_row}:F{last_excel_row})',
                currency_format,
            )
            return {
                "next_start": total_row + 2,
                "reimbursement_total_cell": f"F{total_row + 1}",
            }

        for person_id, person_name in sorted(persons.keys(), key=lambda item: (name_key(item[1]), str(item[0]))):
            base_sheet_name = str(person_name) if str(person_id) == str(person_name) else f"{person_id} - {person_name}"
            sheet_name = _safe_sheet_name(base_sheet_name, used_sheet_names)
            hourly_rows = hourly_events.get((person_id, person_name), [])
            person_turno = turno_events.get((person_id, person_name), {})
            expense_rows = expense_events.get((person_id, person_name), [])
            role = _person_role(hourly_rows, person_turno, expense_rows)
            person_period_start, person_period_end, person_period_text = _person_period(hourly_rows, person_turno, period_end)
            worked_day_count = _count_worked_days(hourly_rows, person_turno, person_period_start, person_period_end)

            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            if str(person_id) == str(person_name):
                worksheet.write(0, 0, f"Name: {person_name}", light_green_text_format)
            else:
                worksheet.write(0, 0, f"Person ID: {person_id}, Name: {person_name}", light_green_text_format)
            worksheet.write(0, 1, "", light_green_text_format)

            if person_period_text:
                worksheet.write(1, 0, f"Period: {person_period_text}")
            worksheet.write_blank(2, 0, None)

            worksheet.set_column("A:A", 24)
            worksheet.set_column("B:D", 14)
            worksheet.set_column("E:E", 10)
            worksheet.set_column("F:G", 12)
            worksheet.set_column("H:H", 45)

            rate_info, found_rate = _get_rate_info(person_id, person_name, rates_dict, rates_by_name)
            if hourly_rows and not found_rate:
                missing_rate_people.add(f"{person_id} - {person_name}")
            hourly_rate = float(rate_info["RATE"])

            current_section_row = 3
            section_hours_cells = []
            section_dollar_cells = []
            section_clean_ranges = []
            expense_reimbursement_cell = None

            # Sections with no rows are omitted entirely.
            if hourly_rows:
                hourly_info = write_hourly_section(worksheet, current_section_row, hourly_rows, hourly_rate)
                section_hours_cells.append(hourly_info["hours_total_cell"])
                section_dollar_cells.append(hourly_info["dollar_total_cell"])
                current_section_row = hourly_info["next_start"]

            for section_header in ["Mango Villas", "Casa Damisela", "MARU"]:
                section_rows = person_turno.get(section_header, [])
                if not section_rows:
                    continue
                section_info = write_location_section(
                    worksheet,
                    current_section_row,
                    section_header,
                    section_rows,
                )
                section_hours_cells.append(section_info["hours_total_cell"])
                section_dollar_cells.append(section_info["dollar_total_cell"])
                section_clean_ranges.append((section_info["clean_start_excel"], section_info["clean_end_excel"]))
                current_section_row = section_info["next_start"]

            other_rows = person_turno.get("Other", [])
            if other_rows:
                other_section_info = write_location_section(
                    worksheet,
                    current_section_row,
                    "Other",
                    other_rows,
                )
                section_hours_cells.append(other_section_info["hours_total_cell"])
                section_dollar_cells.append(other_section_info["dollar_total_cell"])
                current_section_row = other_section_info["next_start"]

            if expense_rows:
                expense_section_info = write_expense_section(worksheet, current_section_row, expense_rows)
                expense_reimbursement_cell = expense_section_info["reimbursement_total_cell"]
                current_section_row = expense_section_info["next_start"]

            summary_header_row = current_section_row
            worksheet.merge_range(summary_header_row, 0, summary_header_row, 6, "Summary", summary_header_format)

            total_hours_block_row = summary_header_row + 1
            worksheet.write(total_hours_block_row, 0, "Total Hours")
            hours_formula = "=" + "+".join(section_hours_cells) if section_hours_cells else "=0"
            worksheet.write_formula(total_hours_block_row, 4, hours_formula)

            extras_row_idx = summary_header_row + 2
            subtotal_row_idx = extras_row_idx + 1
            extras_excel_row = extras_row_idx + 1

            worksheet.write(extras_row_idx, 0, "Extras $")

            start_date = rate_info["START"]
            extra_val = float(rate_info["EXTRA"])
            details_val = rate_info["DETAILS"]
            today = datetime.now()
            show_red = False
            if pd.notna(start_date):
                start_dt = start_date.to_pydatetime() if hasattr(start_date, "to_pydatetime") else start_date
                if (today - start_dt).days < 28 or today.month == 1:
                    allowance_amount = 500
                    show_red = True
                else:
                    allowance_amount = 0
            else:
                allowance_amount = 500

            worksheet.write_number(extras_row_idx, 4, extra_val, currency_format)
            if details_val:
                worksheet.write(extras_row_idx, 5, details_val)

            worksheet.write(subtotal_row_idx, 0, "Subtotal $")
            subtotal_parts = [*section_dollar_cells, f"E{extras_excel_row}"]
            subtotal_formula = "=" + "+".join(subtotal_parts)
            worksheet.write_formula(subtotal_row_idx, 4, subtotal_formula, currency_format)

            next_summary_row = subtotal_row_idx + 1
            expense_reimbursement_excel_row = None
            if expense_reimbursement_cell:
                worksheet.write(next_summary_row, 0, "Expense reimbursements $")
                worksheet.write_formula(next_summary_row, 4, f"={expense_reimbursement_cell}", currency_format)
                expense_reimbursement_excel_row = next_summary_row + 1
                next_summary_row += 1

            exclusion_row_idx = next_summary_row
            worksheet.write(exclusion_row_idx, 0, "No-withholding allowance applied this check $ (max $500/year)")
            if show_red:
                worksheet.write_number(exclusion_row_idx, 4, allowance_amount, soft_red_format)
            else:
                worksheet.write_number(exclusion_row_idx, 4, allowance_amount, currency_format)
            worksheet.data_validation(exclusion_row_idx, 4, exclusion_row_idx, 4, {
                "validate": "decimal",
                "criteria": "between",
                "minimum": 0,
                "maximum": 500,
                "input_title": "Allowance",
                "input_message": "Enter the current-check no-withholding allowance, from $0 to $500.",
                "error_title": "Invalid allowance",
                "error_message": "Enter an amount from $0 to $500.",
            })

            withheld_row_idx = exclusion_row_idx + 1
            total_dollar_idx = withheld_row_idx + 2
            worksheet.write(withheld_row_idx, 0, "10% withheld today $")
            exclusion_excel_row = exclusion_row_idx + 1
            withheld_formula = f"=ROUNDDOWN(MAX(E{subtotal_row_idx + 1}-E{exclusion_excel_row},0)*0.10,2)"
            worksheet.write_formula(withheld_row_idx, 4, withheld_formula, currency_format)

            worksheet.write(total_dollar_idx, 0, "Total $")
            final_total_formula = f"=E{subtotal_row_idx + 1} - E{withheld_row_idx + 1}"
            if expense_reimbursement_excel_row:
                final_total_formula = f"{final_total_formula} + E{expense_reimbursement_excel_row}"
            worksheet.write_formula(total_dollar_idx, 4, final_total_formula, light_green_currency_format)

            review_cell = f"F{total_dollar_idx + 1}"
            worksheet.write_blank(total_dollar_idx, 5, None)
            worksheet.data_validation(total_dollar_idx, 5, total_dollar_idx, 5, {
                "validate": "list",
                "source": ["", "y"],
                "input_title": "Reviewed?",
                "input_message": "Choose 'y' once this sheet is reviewed.",
            })

            sheet_ref = quote_sheetname(sheet_name)
            hours_cell = f"={sheet_ref}!E{total_hours_block_row + 1}"
            total_cell = f"={sheet_ref}!E{total_dollar_idx + 1}"
            withheld_cell = f"={sheet_ref}!E{withheld_row_idx + 1}"
            reviewed_cell = f'=IF({sheet_ref}!{review_cell}="y","y","")'
            section_clean_counts = [
                f'COUNTIF({sheet_ref}!F{start}:F{end},"<>")' for start, end in section_clean_ranges
            ]
            cleans_formula = "=" + "+".join(section_clean_counts) if section_clean_counts else "=0"

            summary_entries.append((
                sheet_name,
                role,
                person_period_text,
                worked_day_count,
                hours_cell,
                cleans_formula,
                total_cell,
                withheld_cell,
                reviewed_cell,
            ))

        summary_sheet.write(0, 0, "Person", summary_header_format)
        summary_sheet.write(0, 1, "Role", summary_header_format)
        summary_sheet.write(0, 2, "Period", summary_header_format)
        summary_sheet.write(0, 3, "Total Days", summary_header_format)
        summary_sheet.write(0, 4, "Total Hours", summary_header_format)
        summary_sheet.write(0, 5, "Total Cleans", summary_header_format)
        summary_sheet.write(0, 6, "Total $", summary_header_format)
        summary_sheet.write(0, 7, "Withheld $", summary_header_format)
        summary_sheet.write(0, 8, "Pay/Hour", summary_header_format)
        summary_sheet.write(0, 9, "Pay/Job", summary_header_format)
        summary_sheet.write(0, 10, "Reviewed", summary_header_format)

        for idx, (label, role, entry_period, total_days, hours_ref, cleans_ref, total_ref, withheld_ref, reviewed_ref) in enumerate(summary_entries, start=1):
            excel_row = idx + 1
            summary_sheet.write(idx, 0, label)
            summary_sheet.write(idx, 1, role)
            summary_sheet.write(idx, 2, entry_period)
            summary_sheet.write_number(idx, 3, total_days)
            summary_sheet.write_formula(idx, 4, hours_ref)
            summary_sheet.write_formula(idx, 5, cleans_ref)
            summary_sheet.write_formula(idx, 6, total_ref, currency_format)
            summary_sheet.write_formula(idx, 7, withheld_ref, currency_format)
            summary_sheet.write_formula(idx, 8, f'=IFERROR(G{excel_row}/E{excel_row},"")', currency_format)
            summary_sheet.write_formula(idx, 9, f'=IFERROR(G{excel_row}/F{excel_row},"")', currency_format)
            summary_sheet.write_formula(idx, 10, reviewed_ref)

        if summary_entries:
            total_row = len(summary_entries) + 1
            summary_sheet.write(total_row, 0, "All sheets total", summary_header_format)
            summary_sheet.write_formula(total_row, 3, f"=SUM(D2:D{total_row})")
            summary_sheet.write_formula(total_row, 4, f"=SUM(E2:E{total_row})")
            summary_sheet.write_formula(total_row, 5, f"=SUM(F2:F{total_row})")
            summary_sheet.write_formula(total_row, 6, f"=SUM(G2:G{total_row})", light_green_currency_format)
            summary_sheet.write_formula(total_row, 7, f"=SUM(H2:H{total_row})", currency_format)
        else:
            warnings.append("No usable payroll rows were found; workbook contains only the Summary sheet.")

        all_labels = ([label for label, *_ in summary_entries] + ["All sheets total"]) if summary_entries else ["No payroll rows"]
        summary_sheet.set_column(0, 0, max(len(s) for s in all_labels) + 2)
        summary_sheet.set_column(1, 1, 15)
        summary_sheet.set_column(2, 2, 24)
        summary_sheet.set_column(3, 3, 11)
        summary_sheet.set_column(4, 4, 13)
        summary_sheet.set_column(5, 5, 14)
        summary_sheet.set_column(6, 8, 12)
        summary_sheet.set_column(9, 9, 11)
        summary_sheet.set_column(10, 10, 10)

    for label in sorted(missing_rate_people):
        warnings.append(f"No hourly rate found for {label}; hourly pay defaults to $0.")

    message = f"Excel file '{output_excel}' created successfully."
    return message, warnings


def _run_cli(argv):
    if any(arg.startswith("-") for arg in argv):
        parser = argparse.ArgumentParser(description="Export Optihome payroll timesheets to Excel.")
        parser.add_argument("--output", "-o", required=True, help="Output Excel workbook path.")
        parser.add_argument("--time", help="NGTecoTime timeclock CSV.")
        parser.add_argument("--turno", help="Turno cleaning report CSV.")
        parser.add_argument("--notion", help="Notion contractor timesheet CSV.")
        parser.add_argument("--expenses", help="Notion expenses CSV.")
        parser.add_argument("--rates", help="Employee rates CSV.")
        args = parser.parse_args(argv)
        if not args.time and not args.turno and not args.notion and not args.expenses:
            parser.error("At least one of --time, --turno, --notion, or --expenses is required.")
        return process_timesheet(
            args.time,
            args.output,
            args.turno,
            rates_csv=args.rates,
            notion_csv=args.notion,
            expenses_csv=args.expenses,
        )

    if len(argv) < 2 or len(argv) > 3:
        raise ValueError(
            "Usage: python export-timesheet.py <output_excel> <timeclock_csv> [turno_csv]\n"
            "   or: python export-timesheet.py --output out.xlsx [--time time.csv] [--turno turno.csv] "
            "[--notion notion.csv] [--expenses expenses.csv] [--rates rates.csv]"
        )

    output_excel = argv[0]
    input_csv = argv[1]
    turno_csv = argv[2] if len(argv) == 3 else None
    return process_timesheet(input_csv, output_excel, turno_csv)


if __name__ == "__main__":
    try:
        message, warn_list = _run_cli(sys.argv[1:])
        for warning in warn_list:
            print(f"Warning: {warning}")
        print(message)
    except (FileNotFoundError, ValueError) as exc:
        print(exc)
        sys.exit(1)
