from PyPDF2 import PdfReader
from datetime import datetime
import re

#Self Guided Examples
# example_form = '.\Forms\CompletedBookingForms\Ref. 134_23; Jones, SWCP (Cape Cornwall to Plymouth).pdf'
#example_form = '.\Forms\CompletedBookingForms\Ref. 136_23; Brigham, SWCP (Fowey to Bantham).pdf'
example_form = '.\Forms\CompletedBookingForms\Ref. 029_23; Livingston, South Cornwall (Manopoli party).pdf'



#Day Pack Examples
#example_form = '.\Forms\CompletedBookingForms\Ref. 034_23; Harris, SW5_23.pdf'
def pdfRipper(formLoopItteration, workbookLocation):
    LoopItteration = formLoopItteration
    BookingRef = ""
    BookingDate = ""
    LeadGuest = ""
    Pax = ""
    DepDate = ""
    NumberOfNights = ""
    TotalValue = ""
    Deposit = ""
    OutstandingBalance = ""
    DueDate = ""
    PdfData = []
    reader = PdfReader(example_form)

    #string patterns
    referenceNumberPattern = r"\d{3}/\d{2}"
    bookingDatePattern = r"(\d{1,2}\s*(?:January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{4})"
    firstNamePattern = r"Dear\s+(\w+)"
    lastNamePattern = r";(.*?),"
    numberOfGuestsPattern = r"for (\w+)(?:s?) (?:guest|guests)"
    partyOfSizePattern = r"in a party of (one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty)"
    tourCostPattern = r"tour cost is (\S+)"
    depositPattern = r"deposit of (\S+)"
    paidInFullPattern = r"will be £NIL"
    dueByPattern = r"due\s*by\s*(\d{1,2}\s*(?:January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{4})"
    selfGuidedDepDatePattern = r"\b(\d{1,2}\s*(?:January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{4})\b"#
    selfGuidedNumberOfNightsPattern = r"Number of nights:\s*(\w+)"
    dayPackArrivalDataPattern = r"\b(\d{1,2}-\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})\b"


    #mapping of text numbers to digits
    numberString_to_int = {
        'one': '1',
        'two': '2',
        'three': '3',
        'four': '4',
        'five': '5',
        'six': '6',
        'seven': '7',
        'eight': '8',
        'nine': '9',
        'ten': '10',
        'eleven': '11',
        'twelve': '12',
        'thirteen': '13',
        'fourteen': '14',
        'fifteen': '15',
        'sixteen': '16',
        'seventeen': '17',
        'eighteen': '18',
        'nineteen': '19',
        'twenty': '20'
    }
    #Rips the string from PDF
    orderFormString = " ".join(page.extract_text() for page in reader.pages)

    #Reference Number Logic
    match = re.search(referenceNumberPattern, orderFormString)
    if match:
        reference_number = match.group()                                                                              
        BookingRef = reference_number
    else:
        BookingRef = "missing"

    #Booking Date Logic
    match = re.search(bookingDatePattern, orderFormString)
    if match:
        booking_date_str = match.group(1)
        booking_date = datetime.strptime(booking_date_str, "%d %B %Y")
        booking_date_formatted = booking_date.strftime("%d/%m/%Y")
        BookingDate = booking_date_formatted
    else:
        BookingDate = "missing"

    #Name Logic
    match = re.search(firstNamePattern, orderFormString)
    if match:
        leadGuestString = match.group(1)
        firstName = leadGuestString.split()[0]
        firstName = firstName.replace(" ", "")
    else:
        firstName = "missing"
    match = re.search(lastNamePattern, example_form)
    if match:
        lastName = match.group(1)
        lastName = lastName.replace(" ", "")
    else:
        lastName = "missing"
    LeadGuest = firstName + " " + lastName

    #Party Size Logic
    match = re.search(partyOfSizePattern, orderFormString)
    if match:
        party_size_word = match.group(1)
        party_size_number = numberString_to_int.get(party_size_word)
        Pax = party_size_number
    else:
        match = re.search(numberOfGuestsPattern, orderFormString)
        if match:
            text_guests = match.group(1).lower()
            num_guests = numberString_to_int.get(text_guests)
            Pax = num_guests
        else:
            Pax = "missing"
   
    #Tour Cost Logic
    match = re.search(tourCostPattern, orderFormString)
    if match:
        tour_cost = match.group(1)
        tour_cost = tour_cost.replace(',','')
        TotalValue = tour_cost
    else:
        TotalValue = "missing"

    #Deposit Logic
    match = re.search(depositPattern, orderFormString)
    if match:
        deposit_paid = match.group(1)
        Deposit = deposit_paid
    else:
        match = re.search(paidInFullPattern, orderFormString)
        if match:
            Deposit = tour_cost
        else:
            Deposit = "missing"
    
    #Outstanding Balance Logic
    tour_cost_as_int = tour_cost.replace('£', '')
    tour_cost_as_int = int(tour_cost_as_int);
    deposit_paid_as_int = deposit_paid.replace('£', '')
    deposit_paid_as_int = int(deposit_paid_as_int);
    outstandingBalance = tour_cost_as_int - deposit_paid_as_int
    outstandingBalance = f"£{outstandingBalance}"
    OutstandingBalance = outstandingBalance

    #Due Date logic
    match = re.search(dueByPattern, orderFormString)
    if match:
        due_by_date_str = match.group(1)
        # Convert the due by date string to datetime object
        due_by_date = datetime.strptime(due_by_date_str, "%d %B %Y")
        # Convert the datetime object to the desired date format
        due_by_date_formatted = due_by_date.strftime("%d/%m/%Y")
        DueDate = due_by_date_formatted
    elif(outstandingBalance == "£0"):
        DueDate = BookingDate
    else:
        DueDate = "missing"
    #Check Wether SelfGuided or Day Pacl
    selfGuided = "has been reserved for you" in orderFormString;
    dayPack = "if the tour leader considers that you are poorly equipped" in orderFormString;

    #SelfGuided specific fields
    if(selfGuided and not dayPack):
        #SelfGuided Dep Date
        matches = re.findall(selfGuidedDepDatePattern, orderFormString)
        if len(matches) >= 2:
            arrival_date = matches[1]
            try:
                arrival_date = datetime.strptime(arrival_date, "%d %B %Y").strftime("%d/%m/%Y")
                DepDate = arrival_date
            except ValueError:
                arrival_date = arrival_date.replace(" ", "")
                DepDate = arrival_date
        else:
            DepDate = "missing"
        
        #SelfGuided Number of Nights
        matches = re.findall(selfGuidedNumberOfNightsPattern, orderFormString)
        if matches:
            number_text = matches[0]
            number_text = number_text.lower()
            try:
                numberOfNights = numberString_to_int.get(number_text)
                NumberOfNights = numberOfNights
            except ValueError:
                NumberOfNights = 0
        else:
            NumberOfNights = 0
    
    #Daypack specific fields
    if(dayPack and not selfGuided):

        #Daypack Depdate        
        match = re.search(dayPackArrivalDataPattern, orderFormString)
        if match:
            variable_value = match.group(1)
            # Split the variable value into day range, month, and year
            day_range, month, year = variable_value.split()
            # Extract the start day from the day range
            start_day = day_range.split("-")[0]
            # Create the formatted date string
            formatted_date = f"{start_day} {month} {year}"
            date_object = datetime.strptime(formatted_date, "%d %B %Y")
            formatted_date = date_object.strftime("%d/%m/%Y")
            DepDate = formatted_date
        else:
            DepDate = "missing"
        
        #Daypack Number of Walks(Nights)
        NumberOfNights = orderFormString.count("mile")

    if(dayPack and selfGuided):
        print("Order Form Not Found")

    if(not dayPack and not selfGuided):
        print("Order Form Not Found")
    
    PdfData.append(BookingRef)
    PdfData.append(BookingDate)
    PdfData.append(LeadGuest)
    PdfData.append(Pax)
    PdfData.append(DepDate)
    PdfData.append(NumberOfNights)
    PdfData.append(TotalValue)
    PdfData.append(Deposit)
    PdfData.append(OutstandingBalance)
    PdfData.append(DueDate)
    return PdfData


