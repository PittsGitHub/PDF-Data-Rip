import os
from openpyxl import Workbook
from datetime import datetime


def newBookingFormExport():

    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Booking ref. "
    sheet["B1"] = "Booking date"
    sheet["C1"] = "Lead Guest"
    sheet["D1"] = "Pax"
    sheet["E1"] = "Dep' date"
    sheet["F1"] = "No. nights"
    sheet["G1"] = "Total value"
    sheet["H1"] = "Deposit"
    sheet["I1"] = "Outstanding balance"
    sheet["J1"] = "Due date"

    current_datetime = datetime.now()
    datetime_str = current_datetime.strftime("%d-%m-%y_%H%M%S")

    file_name = f"{datetime_str}-BookingFormData.xlsx"
    folder_path = ".\excel_exports"

    file_path = os.path.join(folder_path, file_name)
    workbook.save(file_path)

    return file_path
