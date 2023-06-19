import services.excelHelpers as excel
import services.pdfDataRipper as ripper

#Create the blank excel template
workbookLocation = excel.newBookingFormExport()
formLoopItteration = 2

pdfRipped = ripper.pdfRipper(formLoopItteration, workbookLocation)
itdo = excel.writeOrderFormData(workbookLocation, formLoopItteration, pdfRipped)

print(itdo)