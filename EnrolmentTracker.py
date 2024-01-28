#---------------------------------------------------------------------------------------------------------------
# Program: excelEnrolmentTracker.py
# Created by: Lauren McMullen 
#---------------------------------------------------------------------------------------------------------------
#  Purpose:
#   This script is for use in the working environment of a marketing assistant. I am tasked with comparing thousands of enrolments at a university to 
#   compare enrolments to webinar attendees. This script will eliminate hours of manual work by storing the emails of students and the emails of webinar attendees 
#   then comapring the two lists and returning list of students that have attended a webinar and enrolled in a class. 
#   The code is designed so that a non-coder in a marketing role can interact and run the code with relative ease by responding to input prompts.
#---------------------------------------------------------------------------------------------------------------

# Import library to enable excel file interaction
import openpyxl 


#Store the emails in the enrolment report
def storeEnrolmentEmails():
    emails_enrolment = []
    wb = openpyxl.load_workbook("") # ATTN INPUT ENROLMENT REPORT .xlsx FILE
    sheet = wb.get_sheet_by_name("Sheet1") # ATTN INPUT SHEET NAME
    for i in range(2,23): # ATTEN SPECIFY MIN/MAX
        cell = "B" + str(i) #ATTN CONFIRM CELL VALUES
        emails_enrolment.append(sheet[cell].value)
    return emails_enrolment
    

# Store the emails of all webinar registrants
def storeWebRegistrants():
    emails_webinar = []
    wb = openpyxl.load_workbook("") # ATTN INPUT WEBINAR ATTENDEE REPORT .xlsx FILE
    sheet = wb.get_sheet_by_name("Sheet1") # ATTN INPUT SHEET NAME
    for i in range(2,23): # ATTN SPECIFY MIN/MAX
        cell = "B" + str(i) #ATTN CONFIRM CELL VALUES
        emails_webinar.append(sheet[cell].value)
    return emails_webinar

# Output a new Excel spreadsheet with all found repeated values 
def checkEmailMatch():
    emails_enrolment = storeEnrolmentEmails
    emails_webinar = storeWebRegistrants
    matchingList = []
    for a in emails_enrolment:
        for b in emails_webinar:
            if emails_enrolment[a] == emails_webinar[b]:
                matchingList.append(emails_enrolment[a])
    return matchingList

print(checkEmailMatch)