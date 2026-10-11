from pathlib import Path

from datetime import datetime, date

import math

import openpyxl



# ---------------- CONFIGURATION ----------------

EXCEL_FILE = Path(__file__).resolve().parent / "12607818.xlsx"

SHEET_NAME = "Daily Log"

HEADER_ROW = 5

DATA_START_ROW = 6

MINUTES_PER_DAY = 1440



# Expected tracking duration confirmed by user: 20 days.

EXPECTED_DAYS = 36



HEADERS = [

    "Date", "Sleep (min)", "Fitness (min)", "Study (min)", "Coding (min)",

    "Class (min)", "Classes Attended", "Other Activities (min)",

    "Total Tracked (min)", "Free/Unaccounted (min)", "Day's Feeling",

    "Satisfaction Level", "Energy Level", "Notes"

]

DURATIONS = ["sleep", "fitness", "study", "coding", "class", "other"]

RATINGS = {

    "feeling": {"excellent", "good", "neutral", "low", "stressed"},

    "satisfaction": {"very satisfied", "satisfied", "neutral", "unsatisfied", "very unsatisfied"},

    "energy": {"high", "medium", "low"}

}

FEELING_SCORE = {"excellent": 5, "good": 4, "neutral": 3, "low": 2, "stressed": 1}

SATISFACTION_SCORE = {"very satisfied": 5, "satisfied": 4, "neutral": 3, "unsatisfied": 2, "very unsatisfied": 1}

ENERGY_SCORE = {"high": 3, "medium": 2, "low": 1}





def is_number(value):

    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)





# ---------------- READ WORKBOOK ----------------

def read_excel_file():

    if not EXCEL_FILE.is_file():

        print(f"ERROR: Excel file not found: {EXCEL_FILE}")

        print("Keep 12607818.xlsx in the same folder as this Python file.")

        return None



    workbook = None

    try:

        workbook = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)

        if SHEET_NAME not in workbook.sheetnames:

            print(f"ERROR: Worksheet '{SHEET_NAME}' not found.")

            print("Available sheets:", workbook.sheetnames)

            return None



        sheet = workbook[SHEET_NAME]

        actual = [sheet.cell(HEADER_ROW, c).value for c in range(1, len(HEADERS) + 1)]

        if actual != HEADERS:

            print(f"ERROR: Headers in row {HEADER_ROW} do not match.")

            print("Expected:", HEADERS)

            print("Found:   ", actual)

            return None



        records = []

        for row in range(DATA_START_ROW, sheet.max_row + 1):

            v = [sheet.cell(row, c).value for c in range(1, len(HEADERS) + 1)]

            # Ignore empty rows and formula-only template rows.

            if not any(x is not None for x in v[:8] + v[10:]):

                continue

            r = {

                "excel_row": row, "date": v[0], "sleep": v[1], "fitness": v[2],

                "study": v[3], "coding": v[4], "class": v[5],

                "classes_attended": v[6], "other": v[7],

                "feeling": v[10], "satisfaction": v[11], "energy": v[12],

                "notes": v[13]

            }

            if all(is_number(r[k]) for k in DURATIONS):

                r["total_tracked"] = sum(r[k] for k in DURATIONS)

                r["free"] = MINUTES_PER_DAY - r["total_tracked"]

            else:

                r["total_tracked"] = r["free"] = None

            records.append(r)



        print(f"Workbook loaded: {EXCEL_FILE.name}")

        print("Daily records found:", len(records))

        return records

    except PermissionError:

        print(f"ERROR: Permission denied: {EXCEL_FILE}")

        print("Close the workbook in Excel and wait for OneDrive sync.")

        return None

    except Exception as error:

        print(f"ERROR while reading Excel: {error}")

        return None

    finally:

        if workbook is not None:

            workbook.close()





# ---------------- VALIDATE DATA ----------------

def validate_data(records):

    valid, invalid, seen_dates = [], [], set()

    for r in records:

        errors = []

        d = r["date"]

        if not isinstance(d, (datetime, date)):

            errors.append("Date is missing or invalid")

        else:

            d = d.date() if isinstance(d, datetime) else d

            if d in seen_dates:

                errors.append("Duplicate date")

            else:

                seen_dates.add(d)



        for k in DURATIONS:

            value = r[k]

            if not is_number(value):

                errors.append(f"{k} must be numeric")

            elif value < 0 or value > MINUTES_PER_DAY:

                errors.append(f"{k} must be between 0 and 1440 minutes")



        if all(is_number(r[k]) for k in DURATIONS):

            if sum(r[k] for k in DURATIONS) > MINUTES_PER_DAY:

                errors.append("Combined durations exceed 1440 minutes")



        attended = r["classes_attended"]

        if not is_number(attended) or attended < 0 or not float(attended).is_integer():

            errors.append("Classes Attended must be a nonnegative whole number")



        for field, options in RATINGS.items():

            value = r[field]

            if not isinstance(value, str) or not value.strip():

                errors.append(f"{field} is missing")

            elif value.strip().lower() not in options:

                errors.append(f"{field} is not an allowed rating value")



        if errors:

            invalid.append((r["excel_row"], errors))

        else:

            r["total_tracked"] = sum(r[k] for k in DURATIONS)

            r["free"] = MINUTES_PER_DAY - r["total_tracked"]

            valid.append(r)



    print("\n========== DATA VALIDATION ==========")

    print("Valid records:", len(valid))

    print("Invalid records:", len(invalid))

    for row, errors in invalid:

        print(f"Excel row {row}: " + "; ".join(errors))

    return valid





# ---------------- INDICES AND AVERAGES ----------------

def average(records, key):

    return sum(r[key] for r in records) / len(records) if records else None





def calculate_time_indices(records):

    if not records:

        return None

    return {

        "TPI": average(records, "coding"),

        "AAI": average(records, "study") + average(records, "class"),

        "PhAI": average(records, "fitness"),

        "SRI": average(records, "sleep"),

        "ABI": average(records, "free"),

        "TUI": average(records, "total_tracked")

    }





def calculate_ei(records):

    if not records:

        return None

    daily = []

    for r in records:

        daily.append((

            FEELING_SCORE[r["feeling"].strip().lower()]

            + SATISFACTION_SCORE[r["satisfaction"].strip().lower()]

            + ENERGY_SCORE[r["energy"].strip().lower()]

        ) / 3)

    return sum(daily) / len(daily)





def calculate_dci(records, expected_days):

    if not isinstance(expected_days, int) or isinstance(expected_days, bool) or expected_days <= 0:

        raise ValueError("EXPECTED_DAYS must be a positive whole number.")

    if len(records) > expected_days:

        raise ValueError("Valid records exceed EXPECTED_DAYS; confirm the period.")

    return len(records) / expected_days * 100





def calculate_pai(indices, ei, dci):
    """
    Calculate Personal Activity Index (PAI) using the formula
    specified in the CAP776 report template.

    PAI = 0.15(TPI) + 0.20(AAI) + 0.15(PhAI)
          + 0.20(SRI) + 0.15(TUI) + 0.10(EI) + 0.05(DCI)
    """
    return (
        0.15 * indices["TPI"]
        + 0.20 * indices["AAI"]
        + 0.15 * indices["PhAI"]
        + 0.20 * indices["SRI"]
        + 0.15 * indices["TUI"]
        + 0.10 * ei
        + 0.05 * dci
    )

# ---------------- REPORT OUTPUT ----------------

def display_averages(records):

    print("\n========== DAILY AVERAGES ==========")

    for label, key in [

        ("Sleep", "sleep"), ("Fitness", "fitness"), ("Study", "study"),

        ("Coding", "coding"), ("Class", "class"), ("Other Activities", "other"),

        ("Total Tracked", "total_tracked"), ("Free / Unaccounted", "free")

    ]:

        print(f"{label}: {average(records, key):.2f} minutes/day")

    print(f"Study + class combined: {average(records, 'study') + average(records, 'class'):.2f} minutes/day")





def display_time_indices(indices):

    print("\n========== TIME-BASED INDICES (RAW AVERAGES) ==========")

    for name, value in indices.items():

        print(f"{name}: {value:.2f} min/day")

    print("These are raw averages, not normalized scores.")





def display_daily_summary(records):

    print("\n========== DAILY SUMMARY ==========")

    for r in records:

        d = r["date"].strftime("%Y-%m-%d")

        academic = r["study"] + r["class"]

        print(f"{d} | Sleep: {r['sleep']} | Fitness: {r['fitness']} | "

              f"Study: {r['study']} | Class: {r['class']} | Coding: {r['coding']} | "

              f"Academic: {academic} | Free: {r['free']} (minutes)")





def relationship_analysis(records):

    if len(records) < 2:

        print("\nRelationship analysis needs at least two valid records.")

        return

    print("\n========== RELATIONSHIP ANALYSIS ==========")



    # Sleep compared with average energy rating.

    sleep_avg = average(records, "sleep")

    print("\n1. Sleep and Energy")

    for label, group in [

        ("At/above-average sleep", [r for r in records if r["sleep"] >= sleep_avg]),

        ("Below-average sleep", [r for r in records if r["sleep"] < sleep_avg])

    ]:

        scores = [ENERGY_SCORE[r["energy"].strip().lower()] for r in group]

        if scores:

            print(f"{label} ({len(group)} days): {sum(scores)/len(scores):.2f} average energy score")

        else:

            print(f"{label}: no records")



    # Study compared with average satisfaction rating.

    study_avg = average(records, "study")

    print("\n2. Study and Satisfaction")

    for label, group in [

        ("At/above-average study", [r for r in records if r["study"] >= study_avg]),

        ("Below-average study", [r for r in records if r["study"] < study_avg])

    ]:

        scores = [SATISFACTION_SCORE[r["satisfaction"].strip().lower()] for r in group]

        if scores:

            print(f"{label} ({len(group)} days): {sum(scores)/len(scores):.2f} average satisfaction score")

        else:

            print(f"{label}: no records")



    print("\n3. Coding and Energy")

    for category in ("high", "medium", "low"):

        group = [r for r in records if r["energy"].strip().lower() == category]

        if group:

            print(f"{category.capitalize()} energy ({len(group)} days): "

                  f"{average(group, 'coding'):.2f} coding min/day")

        else:

            print(f"{category.capitalize()} energy: no records")

    print("\nThese are descriptive patterns, not proof of cause and effect.")





# ---------------- MAIN ----------------

def main():

    records = read_excel_file()

    if records is None:

        return

    valid = validate_data(records)

    if not valid:

        print("No valid records available. Correct the rows listed above.")

        return



    display_averages(valid)

    indices = calculate_time_indices(valid)

    display_time_indices(indices)

    ei = calculate_ei(valid)

    print(f"\nExperience Index (EI): {ei:.2f} / 5")

    print("Confirm that the rating-to-score mapping matches the rubric.")

    display_daily_summary(valid)

    relationship_analysis(valid)



    dates = [r["date"].date() if isinstance(r["date"], datetime) else r["date"] for r in valid]

    span = (max(dates) - min(dates)).days + 1

    print("\n========== DATA COMPLETENESS ==========")

    print(f"Valid logged days: {len(valid)}")

    print(f"Observed date span: {span} calendar days")



    if EXPECTED_DAYS is None:

        print("DCI and PAI pending: confirm the expected tracking days with your teacher.")

        print("Then set EXPECTED_DAYS near the top of this file.")

    else:

        try:

            dci = calculate_dci(valid, EXPECTED_DAYS)

            print(f"DCI: {dci:.2f}%")
            pai = calculate_pai(indices, ei, dci)
            print(f"PAI: {pai:.2f}")
            print("PAI calculated using the predefined weights in the report template.")


        except ValueError as error:

            print(f"Could not calculate DCI/PAI: {error}")



    print("\nProgram finished.")





if __name__ == "__main__":

    main()
