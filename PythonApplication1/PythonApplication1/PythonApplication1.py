#!/usr/bin/python34
#Gets information about names, job titles and emails and send an automated email to all addresses in the list
import openpyxl as pyxl
import smtplib as smt

wb = pyxl.load_workbook('C:\\Users\\LPadon\\Documents\\sample.xlsx')
sendFrom = input("Input your email address:\n")
password = input("Input your password:\n")

smtpObj = smt.SMTP('smtp.gmail.com', 587) #TODO Find out SMTP server
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login(sendFrom, password)


sheetNames = wb.get_sheet_names()

for sheet in sheetNames: #Loop executes the search/email for all 
    activeSheet = wb.get_sheet_by_name(sheet)
    
    #Loop collects the data for one email from a row, then adds it to and sends the email
    for row in range(activeSheet.min_row + 1, activeSheet.max_row + 1): #Assumes there is a title row.
        firstName = activeSheet['A' + str(row)].value
        lastName = activeSheet['C' + str(row)].value
        jobTitle = activeSheet['D' + str(row)].value
        sendTo = activeSheet['E' + str(row)].value #TODO adjust when you get a look at the real sheet
        message = "Dear " + firstName + " " + lastName + ",\nYou are a " + jobTitle + ". Your email address is " + sendTo
        message = 'Subject: %s\n\n%s' % ("This is a subject", message) #TODO Add real subject
        print(message)
        try:
            smtpObj.sendmail(sendFrom, sendTo, message)
        except NameError:
            print("Error: Email not listed on row " + row)
smtpObj.quit()