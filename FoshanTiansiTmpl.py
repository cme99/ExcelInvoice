import re
import ConvertDate as cd
import CreateCustomExcel as ccx
class FoshanTiansiTmpl():
    if __name__ == "__main__":
        print("FoshanTiansi")


def getDataFromFoshanTiansi(df, invoiceHeaderSheet, invoiceItem, i):
    
    sheetRow = i
    indexRow = 0 
    indexCol = 0
    isCommercialInvoiceFound = False
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
    isTotal = False
    startDescriptionIdx = -1
    endDescriptionIdx = -1
    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    poNoIdx = -1
    amountIdx = -1
    poNo = ''
    unit = ''
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        infoFirstCol = str(excelValues[indexRow])
        if re.search(r'COMMERCIAL INVOICE', infoFirstCol) != None and isCommercialInvoiceFound == False:
            isCommercialInvoiceFound = True
            indexRow += 1
        if  re.search(r'(?<=Marks).*', infoFirstCol) != None and isStartDescription == False:
            isStartDescription = True
            i = 0
            while i < len(excelValues[indexRow]):
                value = str(excelValues[indexRow][i])
                if value == 'Item No.':
                    articleNoIdx = i + 1
                if value == 'Qty':
                    qtyIdx = i
                if value == 'Unit Price':
                    unitPriceIdx = i
                if value == 'P.O Number':
                    poNoIdx = i
                if value == 'Amount':
                    amountIdx = i
                i += 1               
            startDescriptionIdx = indexRow + 1
            if startDescriptionIdx > len(excelValues) :
                startDescriptionIdx = len(excelValues)
            elif startDescriptionIdx < 0:
                startDescriptionIdx = 0

        if  re.search(r'(?<=SAY TOTAL:).*', infoFirstCol) != None and isEndDescription == False: 
            isEndDescription = True
            endDescriptionIdx = indexRow - 2
            if endDescriptionIdx > len(excelValues) :
                endDescriptionIdx = len(excelValues)
            elif endDescriptionIdx < 0:
                endDescriptionIdx = 0

        if re.search(r'TOTAL', infoFirstCol) != None and isTotal == False:
            for c in excelValues[indexRow]:
                if c == 'TOTAL':
                    isTotal = True
                total = c
                grandTotal = total #default

        if isCommercialInvoiceFound == True:
            while indexCol < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][indexCol])
                if  re.search(r'(?<=Invoice NO.:).*', info) != None and isInvNoFound == False:
                    invNo = re.findall(r'(?<=Invoice NO.:).*', info)[0].strip()
                    isInvNoFound = True
                    indexRow += 1
                info = str(excelValues[indexRow][indexCol])
                if  re.search(r'(?<=Issue Date:).*', info) != None and isIssueDateFound == False:
                    issueDate = re.findall(r'(?<=Issue Date:).*', info)[0].strip()
                    isIssueDateFound = True
                    indexRow += 1
                indexCol += 1
        indexRow += 1
    # print(invNo + " : " +issueDate)
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    if isStartDescription and isEndDescription:
        # print("Start index : " + str(startDescriptionIdx) + " until index : " + str(endDescriptionIdx))
        while startDescriptionIdx < endDescriptionIdx:
            indexCol = 0
            articleNo = str(excelValues[startDescriptionIdx][articleNoIdx])
            if re.search(r'[0-9]*\.[0-9]*\.?[0-9]*', articleNo) != None:
                # worksheet.write('D' + str(worksheetIndex), str(excelValues[startDescriptionIdx][poNoIdx]))
                if str(excelValues[startDescriptionIdx][poNoIdx]) == 'nan':
                    invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
                else:
                    poNo = str(excelValues[startDescriptionIdx][poNoIdx])
                    invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
                    
                invoiceItemSheet.write('B' + str(worksheetIndex), str(worksheetIndex - 1))
                invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
                if str(excelValues[startDescriptionIdx][qtyIdx + 1]) == 'nan':
                    invoiceItemSheet.write('D' + str(worksheetIndex), unit)
                else:
                    unit = str(excelValues[startDescriptionIdx][qtyIdx + 1])
                    invoiceItemSheet.write('D' + str(worksheetIndex), unit)
                invoiceItemSheet.write('E' + str(worksheetIndex), articleNo) 
                invoiceItemSheet.write('F' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitPriceIdx]))
                invoiceItemSheet.write('G' + str(worksheetIndex), str(round(excelValues[startDescriptionIdx][amountIdx], 2))) 
                # invoiceItemSheet.write('E' + str(worksheetIndex), '') 
                worksheetIndex += 1
            # print(value)
            startDescriptionIdx += 1

    
    invoiceHeaderSheet.write('A' + sheetRow, customerNo)
    invoiceHeaderSheet.write('B' + sheetRow, invNo)
    invoiceHeaderSheet.write('C' + sheetRow, cd.getIssueDate(issueDate))
    invoiceHeaderSheet.write('D' + sheetRow, total)
    invoiceHeaderSheet.write('E' + sheetRow, extendTotal)
    invoiceHeaderSheet.write('F' + sheetRow, grandTotal)
    