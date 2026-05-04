import json
import math
import os
import sys
from datetime import datetime
from tabulate import tabulate
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DATA_FILE = "attendance.json"
THRESHOLD = 75.0


def main():
    print("\nStudent Attendance Manager")
    print("=" * 40)

    while True:
        print("\n1. Add subject")
        print("2. Mark attendance")
        print("3. View attendance summary")
        print("4. Predict classes needed to reach 75%")
        print("5. Simulate: What if I skip next class?")
        print("6. View full log for a subject")
        print("7. Export to Excel")
        print("8. Exit")

        choice = input("\nEnter choice (1-8): ").strip()
        data = load_data(DATA_FILE)

        if choice == "1":
            name = input("Enter subject name: ").strip().title()
            if not name:
                print("Subject name cannot be empty.")
                continue
            if name in data:
                print(f"Subject '{name}' already exists.")
            else:
                data[name] = []
                save_data(DATA_FILE, data)
                print(f"Subject '{name}' added.")

        elif choice == "2":
            if not data:
                print("No subjects found. Please add a subject first.")
                continue
            print("\nSubjects:", ", ".join(data.keys()))
            subject = input("Enter subject: ").strip().title()
            if subject not in data:
                print(f"Subject '{subject}' not found.")
                continue
            status = input("Status (p=Present / a=Absent / l=Late): ").strip().lower()
            status_map = {"p": "Present", "a": "Absent", "l": "Late"}
            if status not in status_map:
                print("Invalid status. Use p, a, or l.")
                continue
            date = input("Date (YYYY-MM-DD) or press Enter for today: ").strip()
            if not date:
                date = datetime.now().strftime("%Y-%m-%d")
            data = mark_attendance(data, subject, status_map[status], date)
            save_data(DATA_FILE, data)
            print(f"Marked {status_map[status]} for '{subject}' on {date}.")

        elif choice == "3":
            if not data:
                print("No subjects found.")
                continue
            summary = get_summary(data)
            overall = get_overall_attendance(summary)

            print(f"\nOverall Attendance: {overall['percentage']:.1f}%"
                  f" ({'Above' if overall['percentage'] >= THRESHOLD else 'Below'} 75%)")
            print(f"Total: {overall['total']}  Attended: {overall['attended']}  Missed: {overall['absent']}")

            table = []
            for subject, info in summary.items():
                pct = info["percentage"]
                status = "OK" if pct >= THRESHOLD else "Below 75%"
                table.append([subject, info["total"], info["present"], info["absent"], info["late"],
                               f"{pct:.1f}%", status])
            print("\n" + tabulate(table,
                                  headers=["Subject", "Total", "Present", "Absent", "Late", "Attendance %", "Status"],
                                  tablefmt="grid"))

        elif choice == "4":
            if not data:
                print("No subjects found.")
                continue
            summary = get_summary(data)
            table = []
            for subject, info in summary.items():
                needed = predict_classes_needed(info["attended"], info["total"])
                msg = "Already above 75%" if needed == 0 else f"Attend {needed} more consecutive class(es)"
                table.append([subject, f"{info['percentage']:.1f}%", msg])
            print("\nClasses Needed to Reach 75%")
            print(tabulate(table, headers=["Subject", "Current %", "What You Need"], tablefmt="grid"))

        elif choice == "5":
            if not data:
                print("No subjects found.")
                continue
            summary = get_summary(data)
            table = []
            for subject, info in summary.items():
                result = simulate_skip(info["attended"], info["total"])
                impact = "Will drop below 75%" if result["drops_below"] else "Still safe"
                table.append([subject, f"{info['percentage']:.1f}%", f"{result['new_percentage']:.1f}%", impact])
            print("\nSimulation: What If I Skip Next Class?")
            print(tabulate(table, headers=["Subject", "Current %", "After Skip %", "Impact"], tablefmt="grid"))

        elif choice == "6":
            if not data:
                print("No subjects found.")
                continue
            print("\nSubjects:", ", ".join(data.keys()))
            subject = input("Enter subject: ").strip().title()
            if subject not in data:
                print(f"Subject '{subject}' not found.")
                continue
            records = filter_by_subject(data, subject)
            if not records:
                print(f"No records found for '{subject}'.")
            else:
                print(f"\nLog for '{subject}':")
                print(tabulate(records, headers=["#", "Date", "Status"], tablefmt="grid"))

        elif choice == "7":
            if not data:
                print("No subjects found.")
                continue
            filename = input("Enter filename (without .xlsx) or press Enter for 'attendance_report': ").strip()
            if not filename:
                filename = "attendance_report"
            summary = get_summary(data)
            filepath = export_to_excel(data, summary, filename + ".xlsx")
            print(f"Exported to '{filepath}'.")

        elif choice == "8":
            print("\nGoodbye.\n")
            sys.exit(0)

        else:
            print("Invalid choice. Please enter 1-8.")


def mark_attendance(data, subject, status, date):
    """
    Add an attendance record for a subject.

    Args:
        data (dict): Full attendance data.
        subject (str): Subject name.
        status (str): 'Present', 'Absent', or 'Late'.
        date (str): Date string in YYYY-MM-DD format.

    Returns:
        dict: Updated attendance data.
    """
    if subject not in data:
        data[subject] = []
    data[subject].append({"date": date, "status": status})
    return data


def get_summary(data):
    """
    Calculate attendance statistics per subject.

    Args:
        data (dict): Full attendance data.

    Returns:
        dict: Per-subject summary with total, present, absent, late, attended, percentage.
              Late counts as attended for percentage purposes.
    """
    summary = {}
    for subject, records in data.items():
        total = len(records)
        present = sum(1 for r in records if r["status"] == "Present")
        absent = sum(1 for r in records if r["status"] == "Absent")
        late = sum(1 for r in records if r["status"] == "Late")
        attended = present + late
        percentage = (attended / total * 100) if total > 0 else 0.0
        summary[subject] = {
            "total": total,
            "present": present,
            "absent": absent,
            "late": late,
            "attended": attended,
            "percentage": percentage,
        }
    return summary


def get_overall_attendance(summary):
    """
    Calculate combined attendance across all subjects.

    Args:
        summary (dict): Output of get_summary().

    Returns:
        dict: total, attended, absent, percentage across all subjects.
    """
    total = sum(s["total"] for s in summary.values())
    attended = sum(s["attended"] for s in summary.values())
    absent = sum(s["absent"] for s in summary.values())
    percentage = (attended / total * 100) if total > 0 else 0.0
    return {"total": total, "attended": attended, "absent": absent, "percentage": percentage}


def predict_classes_needed(attended, total, threshold=75.0):
    """
    Predict how many consecutive classes must be attended to reach the threshold.

    Solves: (attended + x) / (total + x) >= threshold / 100

    Args:
        attended (int): Classes attended so far (present + late).
        total (int): Total classes held so far.
        threshold (float): Target attendance percentage (default 75).

    Returns:
        int: Number of additional classes to attend. 0 if already at/above threshold.
    """
    if total == 0 or (attended / total * 100) >= threshold:
        return 0
    t = threshold / 100
    needed = (t * total - attended) / (1 - t)
    return math.ceil(needed)


def simulate_skip(attended, total, threshold=75.0):
    """
    Simulate the impact of skipping one more class.

    Args:
        attended (int): Classes attended so far.
        total (int): Total classes held so far.
        threshold (float): Target attendance percentage (default 75).

    Returns:
        dict: new_percentage (float) and drops_below (bool).
    """
    new_total = total + 1
    new_percentage = (attended / new_total * 100) if new_total > 0 else 0.0
    was_above = (attended / total * 100) >= threshold if total > 0 else False
    drops_below = was_above and new_percentage < threshold
    return {"new_percentage": new_percentage, "drops_below": drops_below}


def filter_by_subject(data, subject):
    """
    Get all attendance records for a specific subject as a numbered list.

    Args:
        data (dict): Full attendance data.
        subject (str): Subject name.

    Returns:
        list: List of [index, date, status] rows.
    """
    if subject not in data:
        return []
    return [[i + 1, r["date"], r["status"]] for i, r in enumerate(data[subject])]


def export_to_excel(data, summary, filename="attendance_report.xlsx"):
    """
    Export attendance data and summary to a formatted Excel file.

    Args:
        data (dict): Full attendance data (used for the full log sheet).
        summary (dict): Output of get_summary() — precomputed to avoid recomputing internally.
        filename (str): Output filename.

    Returns:
        str: Path of the saved file.
    """
    wb = Workbook()

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F4E79")
    good_fill = PatternFill("solid", start_color="C6EFCE")
    warn_fill = PatternFill("solid", start_color="FFEB9C")
    bad_fill = PatternFill("solid", start_color="FFC7CE")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Sheet 1: Summary
    ws1 = wb.active
    ws1.title = "Summary"

    ws1.merge_cells("A1:H1")
    ws1["A1"] = "Student Attendance Report"
    ws1["A1"].font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    ws1["A1"].alignment = center

    ws1["A2"] = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws1["A2"].font = Font(name="Arial", size=9, italic=True, color="666666")

    headers = ["Subject", "Total Classes", "Present", "Absent", "Late",
               "Attendance %", "Classes Needed (75%)", "Skip Impact"]
    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=4, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    row = 5
    for subject, info in summary.items():
        pct = info["percentage"]
        needed = predict_classes_needed(info["attended"], info["total"])
        skip = simulate_skip(info["attended"], info["total"])
        skip_text = f"{skip['new_percentage']:.1f}% (drops below)" if skip["drops_below"] else f"{skip['new_percentage']:.1f}%"

        values = [
            subject,
            info["total"],
            info["present"],
            info["absent"],
            info["late"],
            f"{pct:.1f}%",
            "Already above 75%" if needed == 0 else f"Attend {needed} more class(es)",
            skip_text,
        ]
        for col, val in enumerate(values, 1):
            cell = ws1.cell(row=row, column=col, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = center
            cell.border = border
            if col == 6:
                cell.fill = good_fill if pct >= 75 else (warn_fill if pct >= 60 else bad_fill)
        row += 1

    overall = get_overall_attendance(summary)
    overall_values = [
        "OVERALL", overall["total"], overall["attended"], overall["absent"],
        "", f"{overall['percentage']:.1f}%", "", "",
    ]
    for col, val in enumerate(overall_values, 1):
        cell = ws1.cell(row=row, column=col, value=val)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for i, w in enumerate([20, 14, 10, 10, 8, 14, 28, 20], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.row_dimensions[4].height = 20

    # Sheet 2: Full Log
    ws2 = wb.create_sheet("Full Log")
    for col, h in enumerate(["Subject", "Date", "Status"], 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    status_colors = {"Present": "C6EFCE", "Absent": "FFC7CE", "Late": "FFEB9C"}
    row = 2
    for subject, records in data.items():
        for r in records:
            for col, val in enumerate([subject, r["date"], r["status"]], 1):
                cell = ws2.cell(row=row, column=col, value=val)
                cell.font = Font(name="Arial", size=10)
                cell.alignment = center
                cell.border = border
                if col == 3:
                    cell.fill = PatternFill("solid", start_color=status_colors.get(r["status"], "FFFFFF"))
            row += 1

    for col, w in zip([1, 2, 3], [20, 14, 12]):
        ws2.column_dimensions[get_column_letter(col)].width = w

    wb.save(filename)
    return filename


def load_data(filepath):
    """
    Load attendance data from a JSON file.

    Args:
        filepath (str): Path to the JSON file.

    Returns:
        dict: Attendance data (subject -> list of records).
    """
    if not os.path.exists(filepath):
        return {}
    with open(filepath) as f:
        return json.load(f)


def save_data(filepath, data):
    """
    Save attendance data to a JSON file.

    Args:
        filepath (str): Path to the JSON file.
        data (dict): Attendance data to save.
    """
    with open(filepath, "w") as f:
        json.dump(data, f, indent=4)


if __name__ == "__main__":
    main()
