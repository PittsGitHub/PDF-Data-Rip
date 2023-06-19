from openpyxl import Workbook

# Create a new workbook
workbook = Workbook()

# Select the active sheet
sheet = workbook.active

# Enter data into cells
sheet["A1"] = "Hello"
sheet["B1"] = "World!"

# Save the workbook
workbook.save("example.xlsx")
