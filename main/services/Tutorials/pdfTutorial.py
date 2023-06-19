from PyPDF2 import PdfReader

# creating a pdf reader object
reader = PdfReader('.\Forms\CompletedBookingForms\Ref. 001_23; Nicholls, Thames Path (source to Oxford).pdf')
  
# printing number of pages in pdf file
print(len(reader.pages))
  
# getting a specific page from the pdf file
page = reader.pages[0]
  
# extracting text from page
text = page.extract_text()
print(text)