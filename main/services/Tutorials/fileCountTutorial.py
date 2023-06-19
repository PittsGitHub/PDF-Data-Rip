import os

folder_path = ".\Forms\CompletedBookingForms"  # Replace with the actual folder path

# Get the list of files in the folder
files = os.listdir(folder_path)

# Initialize a counter variable
pdf_count = 0

# Iterate over each file and check if it ends with ".pdf"
for file_name in files:
    if file_name.endswith(".pdf"):
        pdf_count += 1

# Print the total count of PDF files
print("Total number of PDF files:", pdf_count)