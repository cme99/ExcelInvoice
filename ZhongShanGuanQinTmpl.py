import re
import CreateCustomExcel as ccx
import ConvertDate as cd
class ZhongShanGuanQinTmpl():
    if __name__ == "__main__":
        print("ZhongShanGuanQin")
        
def getDataFromZhongShanGuanQin(df, invoiceHeaderSheet, invoiceItem, i):
    sheetRow = i
    indexRow = 0 
    indexCol = 0
    # isCommercialInvoiceFound = False
    isInvNoFound = False
    isIssueDateFound = False
    invNo = ''
    customerNo = ''
    total = ''
    extendTotal = ''
    grandTotal = ''
    issueDate = ''
    isStartDescription = False
    isEndDescription = False
    # isTotal = False
    startDescriptionIdx = -1
    endDescriptionIdx = -1
    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    # unitRowIdx = -1
    poNoIdx = -1
    amountIdx = -1
    # poNo = ''
    unit = ''
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        infoFirstCol = str(excelValues[indexRow][0])
        # print(infoFirstCol)
        if infoFirstCol == 'INVOICE' and isInvNoFound == False: 
            isInvNoFound = True
            invNo = re.findall(r'(?<=INV NO:).*',str(excelValues[indexRow + 1][0]))[0]
            while indexCol < len(excelValues[indexRow + 1]):
                info = str(excelValues[indexRow + 1][indexCol])
                # print(info)
                if re.search(r'(?<=ISSUE DATE:).*', info) != None and isIssueDateFound == False:
                    issueDate = re.findall(r'(?<=ISSUE DATE:).*', info)[0]
                    isIssueDateFound = True
                    pass
                indexCol += 1
        
        if infoFirstCol == 'Mark & Nos' and isStartDescription == False:
            isStartDescription = True 
            i = 0
            while i < len(excelValues[indexRow]):
                value = str(excelValues[indexRow][i])
                # print("Value : " + value)
                if value == 'HTH No.':
                    articleNoIdx = i
                if value == 'Quantity':
                    qtyIdx = i
                    unit = str(excelValues[indexRow + 1][i])
                if value == 'Unit Price:Usd/Pc':
                    unitPriceIdx = i
                if value == 'P/O NO.':
                    poNoIdx = i
                if re.search(r'Amount', value) != None:
                    amountIdx = i
                i += 1               
            startDescriptionIdx = (indexRow + 2)
            if startDescriptionIdx > len(excelValues) :
                startDescriptionIdx = len(excelValues)
            elif startDescriptionIdx < 0:
                startDescriptionIdx = 0   
        elif re.search(r'(?<=TOTAL VALUE:)', infoFirstCol) != None and isEndDescription == False:    
            # print(infoFirstCol + " : " + str(isEndDescription))
            isEndDescription = True
            total = str(round(excelValues[indexRow -1][amountIdx], 2)) 
            grandTotal = str(round(excelValues[indexRow -1][amountIdx], 2))              
            endDescriptionIdx = (indexRow - 2)
            if endDescriptionIdx > len(excelValues) :
                endDescriptionIdx = len(excelValues)
            elif endDescriptionIdx < 0:
                endDescriptionIdx = 0
                      
        indexRow += 1
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    if isStartDescription and isEndDescription:
        # print("start gain info from description from index : " + str(startDescriptionIdx) + " until index : " + str(endDescriptionIdx))
        
        while startDescriptionIdx <= endDescriptionIdx:
            articleNo = str(excelValues[startDescriptionIdx][articleNoIdx])
            # print("Article No: " + str(excelValues[startDescriptionIdx]))
            if re.search(r'[0-9]*\.[0-9]*\.?[0-9]*', articleNo) != None:
                invoiceItemSheet.write('A' + str(worksheetIndex), str(excelValues[startDescriptionIdx][poNoIdx]))
                invoiceItemSheet.write('B' + str(worksheetIndex), str(worksheetIndex - 1))
                invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
                invoiceItemSheet.write('D' + str(worksheetIndex), unit)
                invoiceItemSheet.write('E' + str(worksheetIndex), articleNo) 
                invoiceItemSheet.write('F' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitPriceIdx]))
                invoiceItemSheet.write('G' + str(worksheetIndex), str(round(excelValues[startDescriptionIdx][amountIdx], 2))) 
                worksheetIndex += 1
            startDescriptionIdx += 1
    invoiceHeaderSheet.write('A' + sheetRow, customerNo)
    invoiceHeaderSheet.write('B' + sheetRow, invNo)
    invoiceHeaderSheet.write('C' + sheetRow, cd.getIssueDateWithoutMonthName(issueDate))
    invoiceHeaderSheet.write('D' + sheetRow, total)
    invoiceHeaderSheet.write('E' + sheetRow, extendTotal)
    invoiceHeaderSheet.write('F' + sheetRow, grandTotal)

