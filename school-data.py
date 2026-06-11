from openpyxl import Workbook

# Create workbook
wb = Workbook()

# School Data Sheet
ws = wb.active
ws.title = "School Data"

# Header
ws.append([
    "School",
    "City",
    "Category",
    "Website",
    "Description",
    "Verification Status",
    "Address",
    "Phone"
])

# School Records
ws.append([
    "The Brearley School",
    "New York",
    "Private Schools",
    "brearley.org",
    "All-girls K–12 on the Upper East Side; consistently ranked among NYC's most academically rigorous independent schools with exceptional Ivy League placement.",
    "Verified",
    "610 East 83rd Street, New York, NY 10028",
    "(212) 744-8582"
])

ws.append([
    "Trinity School",
    "New York",
    "Private Schools",
    "trinityschoolnyc.org",
    "Founded 1709; one of the oldest schools in the US; Episcopal day school K–12 on the Upper West Side; 100% of seniors matriculate to four-year colleges; renowned Classics programme.",
    "Verified",
    "139 West 91st Street, New York, NY 10024",
    "(212) 873-1650"
])

# Missing Data Report Sheet
report = wb.create_sheet("Missing Data Report")

report.append([
    "Location Name",
    "School Name",
    "Missing Field(s)",
    "Reason"
])

report.append([
    "None",
    "None",
    "None",
    "All required information available"
])

# Save Excel File
wb.save("final_school_assignment.xlsx")

print("Excel file created successfully!")
