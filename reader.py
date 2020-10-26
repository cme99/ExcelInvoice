import pandas as pd
import xlsxwriter
import os
import sys
import numpy as np
import re

types = ['FURUAN TRADING CO.,LTD.OF KAIPING CITY', 'ZHONGSHAN GUANGQIN TRADE CO.,LTD.', 'Foshan Tiansi Hardware Co.,Ltd']
def getExcelTemplate(columns, types):
    for column in columns:
        # print (column)
        for type in types:
            # print(type)
            if (column == type):
                return column
    return ''

def getDataFromFuranTrading():
    # arr1 = []
    # iterating over rows using iterrows() function
    index = 0
    startIndex = -1
    endIndex = -1

    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    invoiceNoIdx = -1
    date = ''
    for v in df.values[0]:
        info = ''
        for i, j in df.iterrows():
            info = str(j[index]).strip()
            if (info == 'Date'):
                date = str(df.values[i + 1][index])
                print(date)
            if (j[index] == 'No.'):
                startIndex = i    
            if (j[index] == 'TOTAL QTY:'):
                endIndex = i - 1
            
            if (info == 'Code No.'):
                articleNoIdx = index
            if (info == 'QTY'):
                qtyIdx = index
            if (info == 'Unit price'):
                unitPriceIdx = index
            if (info == 'PO#'):
                invoiceNoIdx = index
        index = index + 1
    # print("INDEX =====INFO======= " + str(articleNoIdx) + " : " + str(qtyIdx) + " : " + str(unitPriceIdx) + " : " + str(invoiceNoIdx))
    # print(str(startIndex) + " : " + str(endIndex))
    startIndex = startIndex + 2
    invoiceNo = ''
    # Workbook() takes one, non-optional, argument  
    # which is the filename that we want to create. 
    workbook = xlsxwriter.Workbook("hello.xlsx") 
    
    # The workbook object is then used to add new  
    # worksheet via the add_worksheet() method. 
    worksheet = workbook.add_worksheet("Result") 
    worksheet.write('A1', 'Code No.')
    worksheet.write('B1', 'QTY')
    worksheet.write('C1', 'Unit price')
    worksheet.write('D1', 'PO#')
    worksheet.write('E1', 'Date')
    worksheet.write('E2', date)
    worksheetIndex = 2
    while startIndex < endIndex:
        index = 0
        arr = []
        # for v in df.values[0]:
            # print("Invoice Description : " + str(df.values[startIndex][index]))
        # arr.append(str(df.values[startIndex][articleNoIdx]))
        # arr.append(str(df.values[startIndex][qtyIdx]))
        # arr.append(str(df.values[startIndex][unitPriceIdx]))
        
        # if str(df.values[startIndex][invoiceNoIdx]) == 'nan':
        #     arr.append(invoiceNo)
        # else:
        #     arr.append(str(df.values[startIndex][invoiceNoIdx]))
        #     invoiceNo = str(df.values[startIndex][invoiceNoIdx])
        
        # index += 1    
        # print("=================")    
        # print(arr)

        startIndex += 1
        
        worksheet.write('A' + str(worksheetIndex), str(df.values[startIndex][articleNoIdx])) 
        worksheet.write('B' + str(worksheetIndex), str(df.values[startIndex][qtyIdx]))
        worksheet.write('C' + str(worksheetIndex), str(df.values[startIndex][unitPriceIdx]))
        if str(df.values[startIndex][invoiceNoIdx]) == 'nan':
            arr.append(invoiceNo)
            worksheet.write('D' + str(worksheetIndex), invoiceNo)
        else:
            worksheet.write('D' + str(worksheetIndex), str(df.values[startIndex][invoiceNoIdx]))
            invoiceNo = str(df.values[startIndex][invoiceNoIdx])
        worksheetIndex += 1    
    # Finally, close the Excel file 
    # via the close() method. 
    workbook.close()  

def getDataFromZhongShanGuanQin():
    invNo = ''
    issueDate = ''
    indexRow = 0 
    indexCol = 0
    isInvFound = False
    isIssueDateFound = False
    isFirstDescription = False
    isEndDescription = False
    indexStartDescription = -1
    indexEndDescription = -1
    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    invoiceNoIdx = -1
    workbook = xlsxwriter.Workbook("Result1.xlsx") 
    
    # The workbook object is then used to add new  
    # worksheet via the add_worksheet() method. 
    worksheet = workbook.add_worksheet("Result") 
    worksheet.write('A1', 'HTH No.')
    worksheet.write('B1', 'Quantity')
    worksheet.write('C1', 'Unit Price:Usd/Pc')
    worksheet.write('D1', 'P/O NO.')
    worksheet.write('E1', 'Date')
    worksheet.write('F1', 'Invoice No')
    worksheetIndex = 2
    while indexRow < len(df.values):
        infoFirstCol = str(df.values[indexRow][0])
        # print(infoFirstCol)
        if infoFirstCol == 'INVOICE' and isInvFound == False: 
            isInvFound = True
            invNo = re.findall(r'(?<=INV NO:).*',str(df.values[indexRow + 1][0]))[0]
            while indexCol < len(df.values[indexRow + 1]):
                info = str(df.values[indexRow + 1][indexCol])
                # print(info)
                if re.search(r'(?<=ISSUE DATE:).*', info) != None and isIssueDateFound == False:
                    issueDate = re.findall(r'(?<=ISSUE DATE:).*', info)[0]
                    isIssueDateFound = True
                    pass
                indexCol += 1
        
        if infoFirstCol == 'Mark & Nos' and isFirstDescription == False:
            isFirstDescription = True 
            i = 0
            while i < len(df.values[indexRow]):
                value = str(df.values[indexRow][i])
                # print("Value : " + value)
                if value == 'HTH No.':
                    articleNoIdx = i
                elif value == 'Quantity':
                    qtyIdx = i
                elif value == 'Unit Price:Usd/Pc':
                    unitPriceIdx = i
                elif value == 'P/O NO.':
                    invoiceNoIdx = i
                i += 1               
            indexStartDescription = (indexRow + 2)
            if indexStartDescription > len(df.values) :
                indexStartDescription = len(df.values)
            elif indexStartDescription < 0:
                indexStartDescription = 0     
            
            
        elif re.search(r'(?<=TOTAL VALUE:)', infoFirstCol) != None and isEndDescription == False:    
            # print(infoFirstCol + " : " + str(isEndDescription))
            isEndDescription = True                
            indexEndDescription = (indexRow - 2)
            if indexEndDescription > len(df.values) :
                indexEndDescription = len(df.values)
            elif indexEndDescription < 0:
                indexEndDescription = 0
                      
        indexRow += 1
    if isFirstDescription and isEndDescription:
        # print("start gain info from description from index : " + str(indexStartDescription) + " until index : " + str(indexEndDescription))
        print("Article No: " + str(articleNoIdx))
        while indexStartDescription <= indexEndDescription:
            articleNo = str(df.values[indexStartDescription][articleNoIdx])
            if re.search(r'[0-9]*\.[0-9]*\.?[0-9]*', articleNo) != None:
                
                worksheet.write('A' + str(worksheetIndex), articleNo) 
                worksheet.write('B' + str(worksheetIndex), str(df.values[indexStartDescription][qtyIdx]))
                worksheet.write('C' + str(worksheetIndex), str(df.values[indexStartDescription][unitPriceIdx]))
                worksheet.write('D' + str(worksheetIndex), str(df.values[indexStartDescription][invoiceNoIdx]))
                # print(articleNo)
                worksheetIndex += 1
            indexStartDescription += 1
    # print(invNo + " : " + issueDate)
    worksheet.write('E2', issueDate)
    worksheet.write('F2', invNo)
    workbook.close()  

def getDataFromFoshanTiansi():
    indexRow = 0 
    indexCol = 0
    isCommercialInvoiceFound = False
    isInvNoFound = False
    isIssueDateFound = False
    invNo = ''
    issueDate = ''
    isStartDescription = False
    isEndDescription = False
    startDescriptionIdx = -1
    endDescriptionIdx = -1
    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    invoiceNoIdx = -1
    invoiceNo = ''
    workbook = xlsxwriter.Workbook("Result2.xlsx") 
    
    # The workbook object is then used to add new  
    # worksheet via the add_worksheet() method. 
    worksheet = workbook.add_worksheet("Result") 
    worksheet.write('A1', 'Article No')
    worksheet.write('B1', 'Quantity')
    worksheet.write('C1', 'Unit Price')
    worksheet.write('D1', 'P/O NO.')
    worksheet.write('E1', 'Date')
    worksheet.write('F1', 'Invoice No')
    worksheetIndex = 2
    while indexRow < len(df.values):
        infoFirstCol = str(df.values[indexRow][0])

        if infoFirstCol == 'COMMERCIAL INVOICE':
            isCommercialInvoiceFound = True
            indexRow += 1
        if  re.search(r'(?<=Marks).*', infoFirstCol) != None and isStartDescription == False:
            print(infoFirstCol)
            isStartDescription = True
            i = 0
            while i < len(df.values[indexRow]):
                value = str(df.values[indexRow][i])
                # print("Value : " + value)
                if value == 'Item No.':
                    articleNoIdx = i + 1
                elif value == 'Qty':
                    qtyIdx = i
                elif value == 'Unit Price':
                    unitPriceIdx = i
                elif value == 'P.O Number':
                    invoiceNoIdx = i
                i += 1               
            startDescriptionIdx = indexRow + 1
            if startDescriptionIdx > len(df.values) :
                startDescriptionIdx = len(df.values)
            elif startDescriptionIdx < 0:
                startDescriptionIdx = 0

        if  re.search(r'(?<=SAY TOTAL:).*', infoFirstCol) != None and isEndDescription == False: 
            isEndDescription = True
            endDescriptionIdx = indexRow - 2
            if endDescriptionIdx > len(df.values) :
                endDescriptionIdx = len(df.values)
            elif endDescriptionIdx < 0:
                endDescriptionIdx = 0

        if isCommercialInvoiceFound == True:
            while indexCol < len(df.values[indexRow]):
                info = str(df.values[indexRow][indexCol])
                if  re.search(r'(?<=Invoice NO.:).*', info) != None and isInvNoFound == False:
                    invNo = re.findall(r'(?<=Invoice NO.:).*', info)[0].strip()
                    isInvNoFound = True
                    indexRow += 1
                info = str(df.values[indexRow][indexCol])
                if  re.search(r'(?<=Issue Date:).*', info) != None and isIssueDateFound == False:
                    issueDate = re.findall(r'(?<=Issue Date:).*', info)[0].strip()
                    isIssueDateFound = True
                    indexRow += 1
                indexCol += 1
        indexRow += 1
    print(invNo + " : " +issueDate)
    if isStartDescription and isEndDescription:
        print("Start index : " + str(startDescriptionIdx) + " until index : " + str(endDescriptionIdx))
        while startDescriptionIdx < endDescriptionIdx:
            indexCol = 0
            articleNo = str(df.values[startDescriptionIdx][articleNoIdx])
            if re.search(r'[0-9]*\.[0-9]*\.?[0-9]*', articleNo) != None:
                
                worksheet.write('A' + str(worksheetIndex), articleNo) 
                worksheet.write('B' + str(worksheetIndex), str(df.values[startDescriptionIdx][qtyIdx]))
                worksheet.write('C' + str(worksheetIndex), str(df.values[startDescriptionIdx][unitPriceIdx]))
                # worksheet.write('D' + str(worksheetIndex), str(df.values[startDescriptionIdx][invoiceNoIdx]))
                if str(df.values[startDescriptionIdx][invoiceNoIdx]) == 'nan':
                    worksheet.write('D' + str(worksheetIndex), invoiceNo)
                else:
                    worksheet.write('D' + str(worksheetIndex), str(df.values[startDescriptionIdx][invoiceNoIdx]))
                    invoiceNo = str(df.values[startDescriptionIdx][invoiceNoIdx])
                # print(articleNo)
                worksheetIndex += 1
            # print(value)
            startDescriptionIdx += 1

    worksheet.write('E2', issueDate)
    worksheet.write('F2', invNo)
    workbook.close()  

strPath = sys.argv[1]
path = r''+strPath
df = pd.read_excel (path) #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'
excelType = getExcelTemplate(df.columns, types)
# print(excelType)
if excelType == 'FURUAN TRADING CO.,LTD.OF KAIPING CITY':
    print ("FURUAN TRADING CO.,LTD.OF KAIPING CITY Template")
    getDataFromFuranTrading()
elif excelType == 'ZHONGSHAN GUANGQIN TRADE CO.,LTD.':
    print ("ZHONGSHAN GUANGQIN TRADE CO.,LTD. Template")
    getDataFromZhongShanGuanQin()
elif excelType == 'Foshan Tiansi Hardware Co.,Ltd':
    print ("Foshan Tiansi Hardware Co.,Ltd Template")
    getDataFromFoshanTiansi()
else:
    print ("No Template Found.")