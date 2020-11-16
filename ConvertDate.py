import datetime
import re

def getMonth(month_name):
    if len(month_name) > 2:
        month_name = month_name[0:3].lower()
        print("Month " + month_name)
        months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        return months.index(month_name) + 1
    return ''

def getIssueDate(issueDate):
    if (issueDate != '') :
        findDates = re.findall(r'\w+', issueDate)
        if (len(findDates) and len(findDates) ==3) :
            month = getMonth(findDates[0])
            date = findDates[1]
            if re.search(r'\d*', date):
                date = re.findall(r'\d*', date)[0]
            issueDate = str(date)+"."+str(month)+"."+str(findDates[2])
    return issueDate

def getIssueDateWithoutMonthName(issueDate):
    if (issueDate != '') :
        findDates = re.findall(r'\d+', issueDate)
        if (len(findDates) and len(findDates) ==3) :
            issueDate = findDates[2]+"."+findDates[1]+"."+findDates[0]
    return issueDate