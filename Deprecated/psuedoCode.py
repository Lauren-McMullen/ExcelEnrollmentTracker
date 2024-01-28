# Purpose:
#   This script is for use in the working environment of a marketing assistant. I am tasked with comparing thousands of enrolments at a university to 
#   compare enrolments to webinar attendees. This script will eliminate hours of manual work by storing the emails of students and the emails of webinar attendees 
#   then comapring the two lists and returning list of students that have attended a webinar and enrolled in a class. 
#   The code is designed so that a non-coder in a marketing role can interact and run the code with relative ease by responding to input prompts. 

# Needed Variables: 
#   enrolmentFileName, enrolmentSheetName, enrolmentMinMax(tuple), enrolmentColumn
#   webinarFileName, webinarSheetName, webinarMinMax(tuple), webinarColumn

# Helper functions:
#   getEnrolementData(): will ask user for input and store variables containing relevant file information for use in the main function
#   getWebinartData(): will ask user for input and store variables containing relevant webinar attendee file information for use in the main function
#   storeEnrolmentData(): Creates the list of users that enrolled in a course. Uses return values from getExcelData, returns a list
#   storeWebinarData(): Creates the list of users that attended a webinar. Uses return values from getExcelData, returns a list
#   checkEmailMatch(): Compares the emails from storeEnrolmentData() & storeWebinarData() to return a list of emails from students that attended a webinar AND enrolled in a course

# Main function:
#   checkEmailMatch(): Will use the helper functions to track crossovers between students who enrolled in the course and attended a webinar. Will print a list of the crossover emails. 

