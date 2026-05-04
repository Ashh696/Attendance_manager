# 📋 Student Attendance Manager

A command-line tool to track, analyze, and export student attendance across multiple subjects — with smart predictions and skip simulations to help you stay above the 75% threshold.

---

## ✨ Features

| Feature | Description |
|---|---|
| **Add Subjects** | Register any number of subjects to track |
| **Mark Attendance** | Log Present, Absent, or Late for any date |
| **Summary View** | Per-subject and overall attendance statistics |
| **Predict Classes Needed** | Know exactly how many classes to attend to reach 75% |
| **Skip Simulator** | See the impact of skipping your next class |
| **Full Log View** | Browse the complete attendance history for any subject |
| **Excel Export** | Export a formatted `.xlsx` report with color-coded data |

---

## 🗂️ Project Structure

```
.
├── attendance.py          # Main application
├── attendance.json        # Auto-generated data file (created on first run)
└── attendance_report.xlsx # Exported report (generated on demand)
```

---

## ⚙️ Requirements

- Python 3.7+
- Dependencies:

```bash
pip install tabulate openpyxl
```

---

## 🚀 Getting Started

```bash
python attendance.py
```

You'll see the main menu:

```
Student Attendance Manager
========================================

1. Add subject
2. Mark attendance
3. View attendance summary
4. Predict classes needed to reach 75%
5. Simulate: What if I skip next class?
6. View full log for a subject
7. Export to Excel
8. Exit
```

---

## 📖 Usage Guide

### 1. Add a Subject
Enter a subject name (e.g., `Mathematics`, `Physics`). Subject names are stored in title case.

### 2. Mark Attendance
Choose a subject, then enter:
- `p` → Present
- `a` → Absent
- `l` → Late *(counts as attended for percentage)*

Enter a date in `YYYY-MM-DD` format, or press **Enter** to use today's date.

### 3. View Attendance Summary
Displays a table with totals, per-subject percentages, and a status indicator (OK / Below 75%).

```
Overall Attendance: 82.3% (Above 75%)
Total: 65  Attended: 54  Missed: 11

+-----------+-------+---------+--------+------+--------------+-----------+
| Subject   | Total | Present | Absent | Late | Attendance % | Status    |
+-----------+-------+---------+--------+------+--------------+-----------+
| Physics   |  20   |   17    |   3    |  0   |    85.0%     | OK        |
| Maths     |  25   |   17    |   7    |  1   |    72.0%     | Below 75% |
+-----------+-------+---------+--------+------+--------------+-----------+
```

### 4. Predict Classes Needed
Uses the formula below to calculate how many consecutive classes you must attend to hit 75%:

```
classes_needed = ceil((0.75 × total − attended) / (1 − 0.75))
```

### 5. Skip Simulator
Shows what your attendance percentage becomes if you miss the next class, and whether it drops below 75%.

### 6. View Full Log
Lists every attendance record for a chosen subject, numbered and sorted by entry order.

### 7. Export to Excel
Saves a `.xlsx` file with two sheets:
- **Summary** — color-coded stats, predictions, and skip impact per subject
- **Full Log** — every record color-coded by status (green = Present, red = Absent, yellow = Late)

---

## 💾 Data Storage

Attendance records are saved locally in `attendance.json`:

```json
{
    "Physics": [
        {"date": "2025-01-10", "status": "Present"},
        {"date": "2025-01-12", "status": "Absent"}
    ]
}
```

The file is created automatically on first use and updated after every action.

---

## 🧮 Core Logic

| Function | Purpose |
|---|---|
| `mark_attendance()` | Appends a record to a subject's list |
| `get_summary()` | Computes totals and percentages per subject |
| `get_overall_attendance()` | Aggregates stats across all subjects |
| `predict_classes_needed()` | Solves for minimum consecutive classes to reach threshold |
| `simulate_skip()` | Calculates new percentage after one absence |
| `filter_by_subject()` | Returns numbered log rows for a subject |
| `export_to_excel()` | Writes formatted `.xlsx` with openpyxl |
| `load_data()` / `save_data()` | JSON persistence helpers |

---

## 📌 Notes

- **Late** is treated as **attended** when computing percentages.
- The attendance threshold is set to **75%** and can be changed by editing the `THRESHOLD` constant at the top of `attendance.py`.
- Dates are not validated for uniqueness — multiple entries for the same date are allowed.

---

## 📄 License

This project is open source. Feel free to modify and distribute it for personal or academic use.
