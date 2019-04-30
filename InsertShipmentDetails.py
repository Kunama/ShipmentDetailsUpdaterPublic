from openpyxl import load_workbook
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import glob
import os
import datetime
import smtplib
import xlrd

fullPath = os.path.abspath("./") + "/"
errorLog = open("errors.txt", "a")
errorEmails = open("errorEmails.txt", "w")
credentials = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scopes='https://www.googleapis.com/auth/spreadsheets')
gc = gspread.authorize(credentials)
# Production Sheet
sheet = gc.open_by_url().sheet1 

for fileName in glob.glob('*.xls*'):
    print(fileName)
    if os.path.splitext(fileName)[1] == ".xlsx":
        ShipmentDetails = load_workbook(filename = fileName)
        #Loads Main worksheet
        ShipmentDetailsValues = ShipmentDetails.active
        #Getting values from file
        InvoiceNumber = ShipmentDetailsValues['B4'].value
        AZNumber = ShipmentDetailsValues['B3'].value
        TotalInvoiceAmount = ShipmentDetailsValues['B6'].value
        BLNumber = ShipmentDetailsValues['B8'].value
        ContainerNumber = ShipmentDetailsValues['B9'].value
        ETA = ShipmentDetailsValues['B20'].value
        SealNumber = ShipmentDetailsValues['B10'].value
        FinalImportingVessel = ShipmentDetailsValues['B16'].value
        Quantity = ShipmentDetailsValues['B7'].value
        ShippedDate = ShipmentDetailsValues['B12'].value

        print(AZNumber)

    if os.path.splitext(fileName)[1] == ".xls":
        ShipmentDetails = xlrd.open_workbook(filename = fileName)
        #Loads Main worksheet
        ShipmentDetailsValues = ShipmentDetails.sheet_by_index(0)
        #Getting values from file
        InvoiceNumber = ShipmentDetailsValues.cell_value(rowx=3,colx=1)
        AZNumber = ShipmentDetailsValues.cell_value(rowx=2,colx=1)
        TotalInvoiceAmount = ShipmentDetailsValues.cell_value(rowx=5,colx=1)
        BLNumber = ShipmentDetailsValues.cell_value(rowx=7,colx=1)
        ContainerNumber = ShipmentDetailsValues.cell_value(rowx=8,colx=1)
        ETA = ShipmentDetailsValues.cell_value(17,1)
        SealNumber = ShipmentDetailsValues.cell_value(rowx=9,colx=1)
        FinalImportingVessel = ShipmentDetailsValues.cell_value(rowx=15,colx=1)
        Quantity = ShipmentDetailsValues.cell_value(rowx=6,colx=1)
        ShippedDate = ShipmentDetailsValues.cell_value(rowx=11,colx=1)
        ShippedDate = datetime.datetime(*xlrd.xldate_as_tuple(ShippedDate, ShipmentDetails.datemode))
        ETA = datetime.datetime(*xlrd.xldate_as_tuple(ETA, ShipmentDetails.datemode))

        print(AZNumber)
    ContainerNumber = ContainerNumber.replace(' ','')
    ContainerNumber = ContainerNumber.replace('-', '')
    BLNumber = BLNumber.replace(' ','')
    BLNumber = BLNumber.replace('-', '')
    AZNumber = AZNumber.replace(' ','')
    try:
        TotalInvoiceAmount = TotalInvoiceAmount.replace('USD', '')
        TotalInvoiceAmount = TotalInvoiceAmount.replace(' ','')
    except AttributeError:
        print("Invoice Total already in INT, can't remove USD")
    try:
        Quantity = Quantity.replace('LBS', '')
        Quantity = Quantity.replace(' ','')
    except AttributeError:
        print("Quantity already in INT, can't remove LBS")

    if isinstance(ETA, datetime.datetime) == False and ETA != None:
        os.rename(fullPath+fileName, fullPath + "Errors/" + fileName)
        print("ETA is not in valid format")
        errorLog.write(AZNumber + " ETA is not in valid format\n")
        errorEmails.write(AZNumber + " ETA is not in valid format\n")
    
    else:
        if ETA == None:
            if os.path.splitext(fileName)[1] == ".xls":
                ETA = ShipmentDetailsValues.cell_value(rowx=18,colx=2)
            elif os.path.splitext(fileName)[1] == ".xlsx":
                ETA = ShipmentDetailsValues['B18'].value
        else:
            ETA = ETA
        now = datetime.datetime.now()
        if ETA>now:
            ETA = ETA.strftime('%m/%d/%Y')
            if isinstance(ShippedDate, datetime.datetime) == False:
                os.rename(fullPath+fileName, fullPath + "Errors/" + fileName)
                print("Shipped Date is not in valid format")
                errorLog.write(AZNumber + " Shipped Date is not in valid format\n")
                errorEmails.write(AZNumber + " Shipped date is not in valid format\n")
            else:
                ExpectedShippedDate = datetime.datetime.now() - datetime.timedelta(days=150)
                if ShippedDate>ExpectedShippedDate:
                    ShippedDate = ShippedDate.strftime('%m/%d/%Y')
                    #Updates cells
                    cell = sheet.find(AZNumber)
                    cell_list = sheet.range('F'+str(cell.row)+':W'+str(cell.row))
                    if(cell_list[0].value == "") and (cell_list[2].value=="") and (cell_list[5].value=="") and (cell_list[6].value=="")  and (cell_list[10].value=="") and (cell_list[11].value=="") and (cell_list[14].value=="") and (cell_list[16].value=="") and (cell_list[17].value==""):
                        cell_list[0].value = InvoiceNumber
                        cell_list[2].value = ETA
                        cell_list[5].value = ShippedDate
                        cell_list[6].value = ContainerNumber
                        cell_list[10].value = TotalInvoiceAmount
                        cell_list[11].value = BLNumber
                        cell_list[14].value = FinalImportingVessel
                        cell_list[16].value = Quantity
                        cell_list[17].value = SealNumber
                        sheet.update_cells(cell_list)
                        errorEmails.write(AZNumber + " has been processed!\n")
                        #Moves files to processed
                        os.rename(fullPath+fileName, fullPath + "Processed/" + fileName)
                    else:
                        os.rename(fullPath+fileName, fullPath + "Errors/" + fileName)
                        print("Something in one of the cells")
                        errorLog.write(AZNumber + " Something in one of the cells\n")
                        errorEmails.write(AZNumber + " Something in one of the cells\n")
                else:
                    os.rename(fullPath+fileName, fullPath + "Errors/" + fileName)
                    print("Shipped Date not right")
                    errorLog.write(AZNumber + " Shipped Date not right\n")
                    errorEmails.write(AZNumber + " Shipped Date not right\n")
        else:
            os.rename(fullPath+fileName, fullPath + "Errors/" + fileName)
            
            print("ETA not in future")
            errorLog.write(AZNumber + " ETA not in future\n")
            errorEmails.write(AZNumber + " ETA not in future")


errorEmails.close()
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login("", "") # Login to email
errorEmailsReadOnly = open("errorEmails.txt", 'r')
msg = errorEmailsReadOnly.read()
server.sendmail("", "", msg) #Send email

server.quit()
