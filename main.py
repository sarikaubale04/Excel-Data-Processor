
import openpyxl

workbook = openpyxl.load_workbook("students.xlsx")
sheet = workbook.active

total_students = sheet.max_row - 1
print("Total students:", total_students)

missing_marks = 0
duplicates = 0
seen = set()

cleaned_sheet = workbook.create_sheet("Cleaned Data")
cleaned_sheet.append(["Name", "Age", "Marks"])

for row in range(2, sheet.max_row + 1):
    name = sheet.cell(row=row, column=1).value
    age = sheet.cell(row=row, column=2).value
    marks = sheet.cell(row=row, column=3).value

    if marks is None:
        missing_marks += 1
        print("Missing marks found in row:", row)

    student = (name, age, marks)

    if student in seen:
        duplicates += 1
        print("Duplicate found:", student)
    else:
        seen.add(student)
        cleaned_sheet.append([name, age, marks])

print("Duplicates removed:", duplicates)
print("Missing marks:", missing_marks)

workbook.save("cleaned_students.xlsx")
print("Cleaned file saved successfully!")
