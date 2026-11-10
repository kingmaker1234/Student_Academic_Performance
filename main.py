"""
==========================================================================
  Student Academic Performance & Parent-Teacher Monitoring System
==========================================================================
  Level   : Beginner
  Type    : Console application (runs in Windows PowerShell)
  Storage : CSV files (created automatically on the first run)
  Modules : csv, os, datetime  (built into Python - nothing to install)

  How to run:
      python main.py
==========================================================================
"""

import csv
import os
from datetime import datetime


# ==========================================================================
# PART 1: SETTINGS - CSV file names and their column headers
# ==========================================================================

# Save the CSV files in the same folder as this main.py file
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

STUDENTS_FILE = os.path.join(BASE_DIR, "students.csv")
TEACHERS_FILE = os.path.join(BASE_DIR, "teachers.csv")
PARENTS_FILE = os.path.join(BASE_DIR, "parents.csv")
GRADES_FILE = os.path.join(BASE_DIR, "grades.csv")
ATTENDANCE_FILE = os.path.join(BASE_DIR, "attendance.csv")
HOMEWORK_FILE = os.path.join(BASE_DIR, "homework.csv")
STUDY_LOGS_FILE = os.path.join(BASE_DIR, "study_logs.csv")
COMMUNICATION_FILE = os.path.join(BASE_DIR, "communication.csv")

# A dictionary: each CSV file -> the list of its column names
FILE_HEADERS = {
    STUDENTS_FILE: ["student_id", "name", "class", "section"],
    TEACHERS_FILE: ["teacher_id", "name", "subject"],
    PARENTS_FILE: ["parent_id", "name", "phone", "student_id"],
    GRADES_FILE: ["student_id", "subject", "marks", "grade", "date"],
    ATTENDANCE_FILE: ["student_id", "date", "status"],
    HOMEWORK_FILE: ["homework_id", "student_id", "subject", "description",
                    "teacher_id", "assigned_date", "status"],
    STUDY_LOGS_FILE: ["student_id", "date", "subject", "hours"],
    COMMUNICATION_FILE: ["date", "student_id", "sender_type", "sender_id",
                         "sender_name", "message"],
}

DATE_FORMAT = "%Y-%m-%d"   # Dates are stored like 2026-11-10
LINE_WIDTH = 75            # Width of the lines printed on screen


# ==========================================================================
# PART 2: CSV FILE HELPERS - create, read, save and add records
# ==========================================================================

def clean_column_name(name):
    """Tidy a column name, for example ' Student ID ' -> 'student_id'."""
    return name.strip().lower().replace(" ", "_")


def check_one_csv_file(file_name, headers):
    """
    Make sure one CSV file exists and has the correct header row.
    Returns True if the file had to be created again or repaired.

    A CSV file can be changed outside this program (for example when it
    is saved in Excel). This function fixes the common problems:
      - Missing or empty file -> a new file with the header row is made.
      - Small header problems -> repaired automatically, data is KEPT.
        (blank lines at the top, spaces, capital letters,
         ';' instead of ',', columns in a different order)
      - Completely wrong columns -> the file is renamed to a backup copy
        and a new empty file is created, so the program does not crash.
    """
    short_name = os.path.basename(file_name)

    # 1. Missing file: create it with just the header row
    if not os.path.exists(file_name):
        save_records(file_name, [])
        return False

    with open(file_name, "r", newline="", encoding="utf-8-sig") as file:
        text = file.read()

    # 2. Keep only the lines that are not empty
    lines = []
    for line in text.splitlines():
        if line.strip() != "":
            lines.append(line)

    if len(lines) == 0:              # empty file: write the header row
        save_records(file_name, [])
        return False

    # 3. Excel in some countries puts ';' between columns instead of ','
    if lines[0].count(";") > lines[0].count(","):
        delimiter = ";"
    else:
        delimiter = ","

    rows = list(csv.reader(lines, delimiter=delimiter))
    found_columns = []
    for name in rows[0]:
        found_columns.append(clean_column_name(name))

    # 4. The file is already correct: nothing to do
    first_line_of_file = text.splitlines()[0]
    if delimiter == "," and rows[0] == headers and first_line_of_file == lines[0]:
        return False

    # 5. All the needed columns are there: repair the file and keep the data
    all_columns_found = True
    for column in headers:
        if column not in found_columns:
            all_columns_found = False

    if all_columns_found:
        records = []
        for row in rows[1:]:
            record = {}
            for column in headers:
                position = found_columns.index(column)
                if position < len(row):
                    record[column] = row[position].strip()
                else:
                    record[column] = ""
            records.append(record)
        save_records(file_name, records)
        print(f"  Repaired {short_name}: the header row was fixed and your data was kept.")
        return True

    # 6. Wrong columns: keep a backup copy, then start a new empty file
    base_name, extension = os.path.splitext(file_name)
    backup_name = base_name + "_backup_" + datetime.now().strftime("%Y%m%d_%H%M%S") + extension
    os.rename(file_name, backup_name)
    save_records(file_name, [])
    print(f"  {short_name} had the wrong columns: {', '.join(rows[0])}")
    print(f"  Expected columns: {', '.join(headers)}")
    print(f"  The old file was saved as {os.path.basename(backup_name)}")
    print(f"  and a new empty {short_name} was created.")
    return True


def check_csv_files():
    """Check every CSV file. Returns True if any file was repaired."""
    anything_repaired = False
    for file_name, headers in FILE_HEADERS.items():
        try:
            if check_one_csv_file(file_name, headers):
                anything_repaired = True
        except PermissionError:
            print(f"  Could not check {os.path.basename(file_name)}. "
                  f"If it is open in Excel, close it and restart the program.")
            anything_repaired = True
    return anything_repaired


def read_records(file_name):
    """Read a CSV file and return a list of dictionaries (one per row)."""
    records = []
    # 'utf-8-sig' also works if Excel added a hidden mark at the file start
    with open(file_name, "r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            records.append(row)
    return records


def save_records(file_name, records):
    """Rewrite the whole CSV file. Used after updating or deleting a record."""
    with open(file_name, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=FILE_HEADERS[file_name])
        writer.writeheader()
        writer.writerows(records)


def append_record(file_name, record):
    """Add one new record (a dictionary) to the end of a CSV file."""
    with open(file_name, "a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=FILE_HEADERS[file_name])
        writer.writerow(record)


def find_record(file_name, id_column, id_value):
    """Return the first record where id_column equals id_value, or None."""
    for record in read_records(file_name):
        if record[id_column] == id_value:
            return record
    return None


# ==========================================================================
# PART 3: INPUT AND DISPLAY HELPERS
# ==========================================================================

def today():
    """Return today's date as text, for example '2026-11-10'."""
    return datetime.now().strftime(DATE_FORMAT)


def get_text(prompt):
    """Keep asking until the user types something (empty input not allowed)."""
    while True:
        value = input(prompt).strip()
        if value != "":
            return value
        print("  This field cannot be empty. Please try again.")


def get_id(prompt):
    """IDs are stored in CAPITAL letters, so 's101' and 'S101' are the same."""
    return get_text(prompt).upper()


def get_number(prompt, minimum, maximum):
    """Ask for a number between minimum and maximum (uses exception handling)."""
    while True:
        try:
            value = float(input(prompt))
        except ValueError:
            print("  That is not a number. Please try again.")
            continue
        if minimum <= value <= maximum:
            return value
        print(f"  Please enter a number from {minimum} to {maximum}.")


def get_date(prompt):
    """Ask for a date (YYYY-MM-DD). Pressing Enter gives today's date."""
    while True:
        value = input(prompt).strip()
        if value == "":
            return today()
        try:
            date = datetime.strptime(value, DATE_FORMAT)   # checks the date is real
            return date.strftime(DATE_FORMAT)              # 2026-1-5 -> 2026-01-05
        except ValueError:
            print("  Invalid date. Use YYYY-MM-DD, for example 2026-11-10.")


def format_number(value):
    """Show 85.0 as '85' and 2.5 as '2.5'."""
    return f"{float(value):g}"


def print_title(title):
    print()
    print("=" * LINE_WIDTH)
    print(title.center(LINE_WIDTH))
    print("=" * LINE_WIDTH)


def print_line():
    print("-" * LINE_WIDTH)


def get_student_names():
    """Return a dictionary of student names, for example {'S101': 'Rahul'}."""
    names = {}
    for student in read_records(STUDENTS_FILE):
        names[student["student_id"]] = student["name"]
    return names


def ask_for_student():
    """Ask for a Student ID. Return the student record, or None if not found."""
    student_id = get_id("Student ID: ")
    student = find_record(STUDENTS_FILE, "student_id", student_id)
    if student is None:
        print(f"  Student {student_id} was not found. Add the student first (option 1).")
    return student


# ==========================================================================
# PART 4: STUDENT MODULE (menu options 1-5)
# ==========================================================================

def add_student():
    print_title("ADD STUDENT")
    student_id = get_id("Student ID: ")
    if find_record(STUDENTS_FILE, "student_id", student_id) is not None:
        print(f"  Student ID {student_id} already exists. Use a different ID.")
        return

    name = get_text("Student Name: ")
    student_class = get_text("Class: ")     # 'class' is a Python keyword, so we use student_class
    section = get_text("Section: ").upper()

    student = {
        "student_id": student_id,
        "name": name,
        "class": student_class,
        "section": section,
    }
    append_record(STUDENTS_FILE, student)
    print(f"\n  Student '{name}' added successfully.")


def print_student_table(students):
    print(f"{'Student ID':<12}{'Name':<25}{'Class':<10}{'Section':<10}")
    print_line()
    for student in students:
        print(f"{student['student_id']:<12}{student['name']:<25}"
              f"{student['class']:<10}{student['section']:<10}")


def view_students():
    print_title("ALL STUDENTS")
    students = read_records(STUDENTS_FILE)
    if len(students) == 0:
        print("  No students found. Add a student first (option 1).")
        return
    print_student_table(students)
    print_line()
    print(f"Total students: {len(students)}")


def search_student():
    print_title("SEARCH STUDENT")
    keyword = get_text("Enter Student ID or Name: ").lower()

    results = []
    for student in read_records(STUDENTS_FILE):
        # Exact match on the ID, or part of the name (so 'rah' finds 'Rahul')
        if keyword == student["student_id"].lower() or keyword in student["name"].lower():
            results.append(student)

    if len(results) == 0:
        print("  No matching student found.")
    else:
        print()
        print_student_table(results)
        print(f"\n{len(results)} student(s) found.")


def update_student():
    print_title("UPDATE STUDENT")
    student_id = get_id("Enter Student ID to update: ")
    students = read_records(STUDENTS_FILE)

    for student in students:
        if student["student_id"] == student_id:
            print("Press Enter to keep the current value.\n")
            new_name = input(f"Student Name [{student['name']}]: ").strip()
            new_class = input(f"Class [{student['class']}]: ").strip()
            new_section = input(f"Section [{student['section']}]: ").strip()

            if new_name != "":
                student["name"] = new_name
            if new_class != "":
                student["class"] = new_class
            if new_section != "":
                student["section"] = new_section.upper()

            save_records(STUDENTS_FILE, students)
            print("\n  Student details updated successfully.")
            return

    print(f"  Student {student_id} was not found.")


def delete_student():
    print_title("DELETE STUDENT")
    student_id = get_id("Enter Student ID to delete: ")
    students = read_records(STUDENTS_FILE)

    for student in students:
        if student["student_id"] == student_id:
            confirm = input(f"Delete {student['name']} ({student_id})? (Y/N): ").strip().upper()
            if confirm == "Y":
                students.remove(student)
                save_records(STUDENTS_FILE, students)
                print("\n  Student deleted successfully.")
                print("  Note: their marks, attendance and other records stay in the other CSV files.")
            else:
                print("  Delete cancelled.")
            return

    print(f"  Student {student_id} was not found.")


# ==========================================================================
# PART 5: TEACHER AND PARENT MODULES (menu options 6-9)
# ==========================================================================

def add_teacher():
    print_title("ADD TEACHER")
    teacher_id = get_id("Teacher ID: ")
    if find_record(TEACHERS_FILE, "teacher_id", teacher_id) is not None:
        print(f"  Teacher ID {teacher_id} already exists. Use a different ID.")
        return

    name = get_text("Teacher Name: ")
    subject = get_text("Subject: ").title()

    append_record(TEACHERS_FILE, {
        "teacher_id": teacher_id,
        "name": name,
        "subject": subject,
    })
    print(f"\n  Teacher '{name}' ({subject}) added successfully.")


def view_teachers():
    print_title("ALL TEACHERS")
    teachers = read_records(TEACHERS_FILE)
    if len(teachers) == 0:
        print("  No teachers found. Add a teacher first (option 6).")
        return

    print(f"{'Teacher ID':<12}{'Name':<25}{'Subject':<20}")
    print_line()
    for teacher in teachers:
        print(f"{teacher['teacher_id']:<12}{teacher['name']:<25}{teacher['subject']:<20}")
    print_line()
    print(f"Total teachers: {len(teachers)}")


def register_parent():
    print_title("REGISTER PARENT")
    parent_id = get_id("Parent ID: ")
    if find_record(PARENTS_FILE, "parent_id", parent_id) is not None:
        print(f"  Parent ID {parent_id} already exists. Use a different ID.")
        return

    name = get_text("Parent Name: ")

    while True:
        phone = get_text("Phone Number: ")
        if phone.isdigit() and len(phone) == 10:
            break
        print("  Phone number must be exactly 10 digits, for example 9876543210.")

    student_id = get_id("Child's Student ID: ")
    student = find_record(STUDENTS_FILE, "student_id", student_id)
    if student is None:
        print(f"  Student {student_id} was not found. Add the student first, then register the parent.")
        return

    # The student_id column is what LINKS (connects) the parent to the child
    append_record(PARENTS_FILE, {
        "parent_id": parent_id,
        "name": name,
        "phone": phone,
        "student_id": student_id,
    })
    print(f"\n  Parent '{name}' registered and linked to {student['name']} ({student_id}).")


def view_parents():
    print_title("ALL PARENTS")
    parents = read_records(PARENTS_FILE)
    if len(parents) == 0:
        print("  No parents found. Register a parent first (option 8).")
        return

    names = get_student_names()
    print(f"{'Parent ID':<11}{'Name':<20}{'Phone':<14}{'Child ID':<10}{'Child Name':<15}")
    print_line()
    for parent in parents:
        child_name = names.get(parent["student_id"], "(deleted)")
        print(f"{parent['parent_id']:<11}{parent['name']:<20}{parent['phone']:<14}"
              f"{parent['student_id']:<10}{child_name:<15}")
    print_line()
    print(f"Total parents: {len(parents)}")


# ==========================================================================
# PART 6: GRADES MODULE (menu options 10-11)
# ==========================================================================

def calculate_grade(marks):
    """Convert marks (0-100) into a letter grade."""
    if marks >= 90:
        return "A+"
    elif marks >= 80:
        return "A"
    elif marks >= 70:
        return "B"
    elif marks >= 60:
        return "C"
    elif marks >= 50:
        return "D"
    else:
        return "F"


def get_student_marks(student_id):
    """Return a list of the grade records for one student."""
    marks_list = []
    for record in read_records(GRADES_FILE):
        if record["student_id"] == student_id:
            marks_list.append(record)
    return marks_list


def calculate_average(marks_list):
    """Average = total marks / number of subjects."""
    if len(marks_list) == 0:
        return 0
    total = 0
    for record in marks_list:
        total += float(record["marks"])
    return total / len(marks_list)


def enter_marks():
    print_title("ENTER STUDENT MARKS")
    student = ask_for_student()
    if student is None:
        return

    subject = get_text("Subject: ").title()
    marks = get_number("Enter marks (0-100): ", 0, 100)
    grade = calculate_grade(marks)

    grades = read_records(GRADES_FILE)

    # If this subject already has marks, update them instead of adding a duplicate
    for record in grades:
        if record["student_id"] == student["student_id"] and record["subject"] == subject:
            record["marks"] = format_number(marks)
            record["grade"] = grade
            record["date"] = today()
            save_records(GRADES_FILE, grades)
            print(f"\n  Marks updated: {subject} = {format_number(marks)} (Grade {grade})")
            return

    append_record(GRADES_FILE, {
        "student_id": student["student_id"],
        "subject": subject,
        "marks": format_number(marks),
        "grade": grade,
        "date": today(),
    })
    print(f"\n  Marks saved: {subject} = {format_number(marks)} (Grade {grade})")


def view_grades():
    print_title("VIEW GRADES")
    student_id = input("Student ID (press Enter to see all students): ").strip().upper()

    if student_id == "":
        grades = read_records(GRADES_FILE)
    else:
        grades = get_student_marks(student_id)

    if len(grades) == 0:
        print("  No marks found.")
        return

    names = get_student_names()
    print(f"\n{'Student ID':<12}{'Name':<20}{'Subject':<20}{'Marks':<8}{'Grade':<6}")
    print_line()
    for record in grades:
        name = names.get(record["student_id"], "(deleted)")
        print(f"{record['student_id']:<12}{name:<20}{record['subject']:<20}"
              f"{record['marks']:<8}{record['grade']:<6}")

    if student_id != "":
        average = calculate_average(grades)
        print_line()
        print(f"Average: {average:.2f}   Overall grade: {calculate_grade(average)}")


# ==========================================================================
# PART 7: ATTENDANCE MODULE (menu options 12-13)
# ==========================================================================

def mark_attendance():
    print_title("MARK ATTENDANCE")
    student = ask_for_student()
    if student is None:
        return

    date = get_date("Date (YYYY-MM-DD, press Enter for today): ")

    while True:
        choice = get_text("Present or Absent (P/A): ").upper()
        if choice == "P":
            status = "Present"
            break
        elif choice == "A":
            status = "Absent"
            break
        print("  Please type P for Present or A for Absent.")

    attendance = read_records(ATTENDANCE_FILE)

    # Only one attendance entry per student per day
    for record in attendance:
        if record["student_id"] == student["student_id"] and record["date"] == date:
            record["status"] = status
            save_records(ATTENDANCE_FILE, attendance)
            print(f"\n  Attendance updated: {student['name']} - {status} on {date}")
            return

    append_record(ATTENDANCE_FILE, {
        "student_id": student["student_id"],
        "date": date,
        "status": status,
    })
    print(f"\n  Attendance saved: {student['name']} - {status} on {date}")


def get_attendance_summary(student_id):
    """Return (total days, present days, attendance percentage) for one student."""
    total_days = 0
    present_days = 0
    for record in read_records(ATTENDANCE_FILE):
        if record["student_id"] == student_id:
            total_days += 1
            if record["status"] == "Present":
                present_days += 1

    if total_days == 0:
        percentage = 0
    else:
        percentage = present_days / total_days * 100
    return total_days, present_days, percentage


def attendance_report():
    print_title("ATTENDANCE REPORT")
    students = read_records(STUDENTS_FILE)
    if len(students) == 0:
        print("  No students found.")
        return

    print(f"{'Student ID':<12}{'Name':<20}{'Days':<8}{'Present':<10}{'Absent':<9}{'Attendance':<10}")
    print_line()
    for student in students:
        total_days, present_days, percentage = get_attendance_summary(student["student_id"])
        absent_days = total_days - present_days
        if total_days == 0:
            percent_text = "No data"
        else:
            percent_text = f"{percentage:.1f}%"
        print(f"{student['student_id']:<12}{student['name']:<20}{total_days:<8}"
              f"{present_days:<10}{absent_days:<9}{percent_text:<10}")


# ==========================================================================
# PART 8: HOMEWORK MODULE (menu options 14-16)
# ==========================================================================

def next_homework_id(homework_list):
    """Create the next homework ID: H1, H2, H3 ..."""
    highest = 0
    for homework in homework_list:
        try:
            number = int(homework["homework_id"][1:])   # 'H12' -> 12
        except ValueError:
            continue
        if number > highest:
            highest = number
    return f"H{highest + 1}"


def get_student_homework(student_id):
    """Return a list of the homework records for one student."""
    result = []
    for homework in read_records(HOMEWORK_FILE):
        if homework["student_id"] == student_id:
            result.append(homework)
    return result


def print_homework_table(homework_list):
    print(f"{'HW ID':<7}{'Student':<10}{'Subject':<15}{'Assigned':<12}{'Status':<11}Description")
    print_line()
    for homework in homework_list:
        print(f"{homework['homework_id']:<7}{homework['student_id']:<10}{homework['subject']:<15}"
              f"{homework['assigned_date']:<12}{homework['status']:<11}{homework['description']}")


def assign_homework():
    print_title("ASSIGN HOMEWORK")
    student = ask_for_student()
    if student is None:
        return

    subject = get_text("Subject: ").title()
    description = get_text("Homework Description: ")
    teacher_id = get_id("Teacher ID: ")
    teacher = find_record(TEACHERS_FILE, "teacher_id", teacher_id)
    if teacher is None:
        print(f"  Teacher {teacher_id} was not found. Add the teacher first (option 6).")
        return

    homework_id = next_homework_id(read_records(HOMEWORK_FILE))

    append_record(HOMEWORK_FILE, {
        "homework_id": homework_id,
        "student_id": student["student_id"],
        "subject": subject,
        "description": description,
        "teacher_id": teacher_id,
        "assigned_date": today(),
        "status": "Pending",
    })
    print(f"\n  Homework {homework_id} assigned to {student['name']} by {teacher['name']}.")


def view_homework():
    print_title("VIEW HOMEWORK")
    student_id = input("Student ID (press Enter to see all homework): ").strip().upper()

    if student_id == "":
        homework_list = read_records(HOMEWORK_FILE)
    else:
        homework_list = get_student_homework(student_id)

    if len(homework_list) == 0:
        print("  No homework found.")
        return
    print()
    print_homework_table(homework_list)


def update_homework_status():
    print_title("UPDATE HOMEWORK STATUS")
    student = ask_for_student()
    if student is None:
        return

    student_homework = get_student_homework(student["student_id"])
    if len(student_homework) == 0:
        print(f"  No homework found for {student['name']}.")
        return

    print()
    print_homework_table(student_homework)
    print()

    homework_id = get_id("Enter Homework ID to update: ")
    homework_list = read_records(HOMEWORK_FILE)

    for homework in homework_list:
        if homework["homework_id"] == homework_id and homework["student_id"] == student["student_id"]:
            print("1. Pending")
            print("2. Completed")
            choice = get_text("Choose new status (1/2): ")
            if choice == "1":
                homework["status"] = "Pending"
            elif choice == "2":
                homework["status"] = "Completed"
            else:
                print("  Invalid choice. Status not changed.")
                return
            save_records(HOMEWORK_FILE, homework_list)
            print(f"\n  Homework {homework_id} is now marked as {homework['status']}.")
            return

    print(f"  Homework {homework_id} was not found for this student.")


# ==========================================================================
# PART 9: STUDY MONITORING MODULE (menu options 17-18)
# ==========================================================================

def get_study_logs(student_id):
    """Return a list of the study log records for one student."""
    logs = []
    for record in read_records(STUDY_LOGS_FILE):
        if record["student_id"] == student_id:
            logs.append(record)
    return logs


def add_study_log():
    print_title("ADD DAILY STUDY LOG")
    student = ask_for_student()
    if student is None:
        return

    subject = get_text("Subject studied: ").title()
    hours = get_number("Study hours: ", 0.25, 24)
    date = get_date("Date (YYYY-MM-DD, press Enter for today): ")

    append_record(STUDY_LOGS_FILE, {
        "student_id": student["student_id"],
        "date": date,
        "subject": subject,
        "hours": format_number(hours),
    })
    print(f"\n  Study log saved: {student['name']} studied {subject} "
          f"for {format_number(hours)} hour(s) on {date}.")


def view_study_report():
    print_title("STUDY REPORT")
    student = ask_for_student()
    if student is None:
        return

    logs = get_study_logs(student["student_id"])
    if len(logs) == 0:
        print(f"  No study logs found for {student['name']}.")
        return

    print(f"\nStudy log for {student['name']} ({student['student_id']})\n")
    print(f"{'Date':<14}{'Subject':<20}{'Hours':<8}")
    print_line()

    total_hours = 0
    hours_by_subject = {}    # dictionary: subject -> total hours
    study_days = []          # list of the different dates studied

    for record in logs:
        hours = float(record["hours"])
        total_hours += hours
        subject = record["subject"]
        hours_by_subject[subject] = hours_by_subject.get(subject, 0) + hours
        if record["date"] not in study_days:
            study_days.append(record["date"])
        print(f"{record['date']:<14}{subject:<20}{format_number(hours):<8}")

    print_line()
    print("Hours by subject:")
    for subject, hours in hours_by_subject.items():
        print(f"  {subject:<20}{format_number(hours)} hour(s)")
    print_line()
    print(f"Total study hours     : {format_number(total_hours)}")
    print(f"Days studied          : {len(study_days)}")
    print(f"Average hours per day : {total_hours / len(study_days):.2f}")


# ==========================================================================
# PART 10: COMMUNICATION MODULE (menu options 19-20)
# ==========================================================================

def get_student_messages(student_id):
    """Return a list of the messages about one student."""
    messages = []
    for record in read_records(COMMUNICATION_FILE):
        if record["student_id"] == student_id:
            messages.append(record)
    return messages


def add_communication():
    print_title("PARENT-TEACHER COMMUNICATION")
    student = ask_for_student()
    if student is None:
        return

    sender_choice = get_text("Who is sending the message? (1 = Teacher, 2 = Parent): ")

    if sender_choice == "1":
        teacher_id = get_id("Teacher ID: ")
        teacher = find_record(TEACHERS_FILE, "teacher_id", teacher_id)
        if teacher is None:
            print(f"  Teacher {teacher_id} was not found.")
            return
        sender_type = "Teacher"
        sender_id = teacher_id
        sender_name = teacher["name"]

    elif sender_choice == "2":
        parent_id = get_id("Parent ID: ")
        parent = find_record(PARENTS_FILE, "parent_id", parent_id)
        if parent is None:
            print(f"  Parent {parent_id} was not found.")
            return
        # A parent can only send feedback about their own child
        if parent["student_id"] != student["student_id"]:
            print(f"  Parent {parent_id} is not linked to {student['name']}.")
            return
        sender_type = "Parent"
        sender_id = parent_id
        sender_name = parent["name"]

    else:
        print("  Invalid choice. Please enter 1 for Teacher or 2 for Parent.")
        return

    message = get_text("Message: ")

    append_record(COMMUNICATION_FILE, {
        "date": today(),
        "student_id": student["student_id"],
        "sender_type": sender_type,
        "sender_id": sender_id,
        "sender_name": sender_name,
        "message": message,
    })
    print(f"\n  Message from {sender_type} {sender_name} saved for {student['name']}.")


def view_communication():
    print_title("VIEW COMMUNICATION")
    student = ask_for_student()
    if student is None:
        return

    messages = get_student_messages(student["student_id"])
    if len(messages) == 0:
        print(f"  No messages found for {student['name']}.")
        return

    print(f"\nMessages about {student['name']} ({student['student_id']})")
    print_line()
    for message in messages:
        print(f"{message['date']}  {message['sender_type']}: "
              f"{message['sender_name']} ({message['sender_id']})")
        print(f"    {message['message']}")
        print()


# ==========================================================================
# PART 11: REPORTS (menu options 21-22)
# ==========================================================================

def student_performance_report():
    print_title("STUDENT PERFORMANCE REPORT")
    student = ask_for_student()
    if student is None:
        return
    student_id = student["student_id"]

    # ---- Student and parent details ----
    print_line()
    print(f"Student ID : {student_id}")
    print(f"Name       : {student['name']}")
    print(f"Class      : {student['class']}   Section: {student['section']}")

    parent_details = []
    for parent in read_records(PARENTS_FILE):
        if parent["student_id"] == student_id:
            parent_details.append(f"{parent['name']} ({parent['phone']})")
    if len(parent_details) == 0:
        print("Parent(s)  : Not registered")
    else:
        print(f"Parent(s)  : {', '.join(parent_details)}")

    # ---- 1) Marks and grades ----
    print("\n1) MARKS AND GRADES")
    print_line()
    marks_list = get_student_marks(student_id)
    if len(marks_list) == 0:
        print("  No marks entered yet.")
    else:
        print(f"{'Subject':<20}{'Marks':<10}{'Grade':<6}")
        total_marks = 0
        for record in marks_list:
            print(f"{record['subject']:<20}{record['marks']:<10}{record['grade']:<6}")
            total_marks += float(record["marks"])
        average = total_marks / len(marks_list)
        print_line()
        print(f"Total marks   : {format_number(total_marks)} / {len(marks_list) * 100}")
        print(f"Average       : {average:.2f}%")
        print(f"Overall grade : {calculate_grade(average)}")

    # ---- 2) Attendance ----
    print("\n2) ATTENDANCE")
    print_line()
    total_days, present_days, percentage = get_attendance_summary(student_id)
    if total_days == 0:
        print("  No attendance recorded yet.")
    else:
        print(f"Days recorded : {total_days}")
        print(f"Present       : {present_days}")
        print(f"Absent        : {total_days - present_days}")
        print(f"Attendance    : {percentage:.1f}%")

    # ---- 3) Homework ----
    print("\n3) HOMEWORK")
    print_line()
    homework_list = get_student_homework(student_id)
    if len(homework_list) == 0:
        print("  No homework assigned yet.")
    else:
        completed = 0
        for homework in homework_list:
            if homework["status"] == "Completed":
                completed += 1
        pending = len(homework_list) - completed
        completion_rate = completed / len(homework_list) * 100
        print(f"Assigned      : {len(homework_list)}")
        print(f"Completed     : {completed}")
        print(f"Pending       : {pending}")
        print(f"Completion    : {completion_rate:.1f}%")

    # ---- 4) Study hours ----
    print("\n4) STUDY HOURS")
    print_line()
    logs = get_study_logs(student_id)
    if len(logs) == 0:
        print("  No study hours recorded yet.")
    else:
        total_hours = 0
        for record in logs:
            total_hours += float(record["hours"])
        print(f"Total study hours : {format_number(total_hours)}")
        print(f"Study sessions    : {len(logs)}")

    # ---- 5) Latest remarks and feedback ----
    print("\n5) LATEST REMARKS AND FEEDBACK")
    print_line()
    messages = get_student_messages(student_id)
    if len(messages) == 0:
        print("  No remarks or feedback yet.")
    else:
        for message in messages[-3:]:          # the last 3 messages
            print(f"{message['date']}  {message['sender_type']} ({message['sender_name']}): "
                  f"{message['message']}")
    print_line()


def get_average_for_sort(item):
    """Used by sort(): tells Python to sort the ranking by the 'average' value."""
    return item["average"]


def class_topper():
    print_title("CLASS TOPPER")
    class_name = input("Enter Class (press Enter for all classes): ").strip()

    ranking = []    # list of dictionaries, one for each student who has marks
    for student in read_records(STUDENTS_FILE):
        if class_name != "" and student["class"].lower() != class_name.lower():
            continue
        marks_list = get_student_marks(student["student_id"])
        if len(marks_list) == 0:
            continue
        ranking.append({
            "student": student,
            "subjects": len(marks_list),
            "average": round(calculate_average(marks_list), 2),
        })

    if len(ranking) == 0:
        print("  No marks found for this class.")
        return

    # Highest average first
    ranking.sort(key=get_average_for_sort, reverse=True)

    print(f"\n{'Rank':<6}{'Student ID':<12}{'Name':<20}{'Class':<8}"
          f"{'Subjects':<10}{'Average':<10}{'Grade'}")
    print_line()

    rank = 0
    previous_average = None
    for position in range(len(ranking)):
        item = ranking[position]
        # Students with the same average share the same rank
        if item["average"] != previous_average:
            rank = position + 1
            previous_average = item["average"]
        student = item["student"]
        class_text = f"{student['class']}-{student['section']}"
        print(f"{rank:<6}{student['student_id']:<12}{student['name']:<20}{class_text:<8}"
              f"{item['subjects']:<10}{item['average']:<10.2f}{calculate_grade(item['average'])}")

    print_line()
    top_average = ranking[0]["average"]
    for item in ranking:
        if item["average"] == top_average:
            student = item["student"]
            print(f"TOPPER: {student['name']} ({student['student_id']}) "
                  f"with an average of {top_average:.2f}%")


# ==========================================================================
# PART 12: MAIN MENU
# ==========================================================================

def show_menu():
    print_title("STUDENT ACADEMIC MONITORING SYSTEM")
    print("MAIN MENU".center(LINE_WIDTH))
    print_line()
    print(" 1.  Add Student")
    print(" 2.  View Students")
    print(" 3.  Search Student")
    print(" 4.  Update Student")
    print(" 5.  Delete Student")
    print(" 6.  Add Teacher")
    print(" 7.  View Teachers")
    print(" 8.  Register Parent")
    print(" 9.  View Parents")
    print(" 10. Enter Student Marks")
    print(" 11. View Grades")
    print(" 12. Mark Attendance")
    print(" 13. Attendance Report")
    print(" 14. Assign Homework")
    print(" 15. View Homework")
    print(" 16. Update Homework Status")
    print(" 17. Add Daily Study Log")
    print(" 18. View Study Report")
    print(" 19. Parent-Teacher Communication")
    print(" 20. View Communication")
    print(" 21. Student Performance Report")
    print(" 22. Class Topper")
    print(" 23. Exit")
    print_line()


def main():
    while True:     # The menu repeats until the user chooses 23 (Exit)
        # Create missing CSV files and repair damaged ones before each menu
        if check_csv_files():
            input("\nPress Enter to continue...")

        show_menu()
        choice = input("Enter your choice (1-23): ").strip()

        try:
            if choice == "1":
                add_student()
            elif choice == "2":
                view_students()
            elif choice == "3":
                search_student()
            elif choice == "4":
                update_student()
            elif choice == "5":
                delete_student()
            elif choice == "6":
                add_teacher()
            elif choice == "7":
                view_teachers()
            elif choice == "8":
                register_parent()
            elif choice == "9":
                view_parents()
            elif choice == "10":
                enter_marks()
            elif choice == "11":
                view_grades()
            elif choice == "12":
                mark_attendance()
            elif choice == "13":
                attendance_report()
            elif choice == "14":
                assign_homework()
            elif choice == "15":
                view_homework()
            elif choice == "16":
                update_homework_status()
            elif choice == "17":
                add_study_log()
            elif choice == "18":
                view_study_report()
            elif choice == "19":
                add_communication()
            elif choice == "20":
                view_communication()
            elif choice == "21":
                student_performance_report()
            elif choice == "22":
                class_topper()
            elif choice == "23":
                print("\nThank you for using the Student Academic Monitoring System. Goodbye!")
                break
            else:
                print("  Invalid choice. Please enter a number from 1 to 23.")

        except PermissionError:
            # On Windows this happens when a CSV file is open in Excel
            print("\n  Could not save the data. If a CSV file is open in Excel, close it and try again.")

        except (KeyError, ValueError, IndexError) as error:
            # A CSV file was changed outside the program and has a wrong column or value
            print("\n  Some data in a CSV file is not in the expected format.")
            print(f"  Details: {type(error).__name__}: {error}")
            print("  Column problems are repaired automatically when you return to the menu.")
            print("  If this message comes back, open the CSV files and look for empty or wrong values.")

        input("\nPress Enter to return to the menu...")


# This line means: run main() only when this file is started directly
# (python main.py), not when it is imported by another Python file.
if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\n\nProgram stopped. Goodbye!")