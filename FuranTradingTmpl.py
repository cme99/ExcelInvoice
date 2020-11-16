import re
import CreateCustomExcel as ccx
import ConvertDate as cd
class FuranTradingTmpl():
    if __name__ == "__main__":
        print("Start Furan Trading Template")

def getDataFromFuranTrading(df, invoiceHeaderSheet, invoiceItem, i):
    sheetRow = i
    indexRow = 0 
    indexCol = 0
    # isInvoiceFound = False
    isInvNoFound = False
    # isIssueDateFound = False
    # isIndexFound = False
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
    poNoIdx = -1
    amountIdx = -1
    indexIdx = -1
    poNo = ''
    unit = ''
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        infoFirstCol = str(excelValues[indexRow])
        # print(infoFirstCol)
        if re.search(r'Ref No.:', infoFirstCol) != None and isInvNoFound == False:
            isInvNoFound = True
            indexCol = 0
            while indexCol < len(excelValues[indexRow]):
                # print(str(excelValues[indexRow][indexCol]))
                info = str(excelValues[indexRow][indexCol])
                if re.search(r'Ref No.:', info) != None:
                    invNo = excelValues[indexRow + 1][indexCol]
                elif info == 'Date':
                    issueDate = excelValues[indexRow + 1][indexCol]
                indexCol += 1
        if re.search(r'Code No.', infoFirstCol) != None and isStartDescription == False:
            indexCol = 0
            startDescriptionIdx = indexRow
            isStartDescription = True
            while indexCol < len(excelValues[indexRow]):
                info =  str(excelValues[indexRow][indexCol])
                # print(info)
                # print(re.search(r'Unit price', info) != None)
                if info == 'No.':
                    indexIdx = indexCol
                elif info == 'Code No.':
                    articleNoIdx = indexCol
                elif re.search(r'QTY', info) != None:
                    qtyIdx = indexCol
                elif re.search(r'Unit price', info) != None:
                    unitPriceIdx = indexCol
                elif re.search(r'PO#', info) != None:
                    poNoIdx = indexCol
                elif re.search(r'AMOUNT', info) != None:
                    amountIdx = indexCol
                indexCol += 1
        if re.search(r'FOB JIANGMEN USD TOTAL AMOUNT:', infoFirstCol) != None and isEndDescription == False:
            endDescriptionIdx = indexRow - 1
            isEndDescription = True
            total = excelValues[indexRow][amountIdx]
            grandTotal = total
        indexRow += 1
    # print(invNo)
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, str(invNo).strip())
    startDescriptionIdx = startDescriptionIdx + 2
    worksheetIndex = 2
    if isStartDescription and isEndDescription:
        # print("Start index : " + str(startDescriptionIdx) + " until index : " + str(endDescriptionIdx))
        
        while startDescriptionIdx < endDescriptionIdx:
            # print(str(articleNoIdx))
            # print(str(excelValues[startDescriptionIdx][unitPriceIdx]))
            if str(excelValues[startDescriptionIdx][poNoIdx]) == 'nan':
                # arr.append(poNo)
                invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
            else:
                invoiceItemSheet.write('A' + str(worksheetIndex), str(excelValues[startDescriptionIdx][poNoIdx]))
                poNo = str(excelValues[startDescriptionIdx][poNoIdx])
            invoiceItemSheet.write('B' + str(worksheetIndex), str(excelValues[startDescriptionIdx][indexIdx]))
            invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
            invoiceItemSheet.write('D' + str(worksheetIndex), unit)
            invoiceItemSheet.write('E' + str(worksheetIndex), str(excelValues[startDescriptionIdx][articleNoIdx])) 
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
