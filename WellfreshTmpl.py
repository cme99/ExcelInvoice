import re
import CreateCustomExcel as ccx
import ConvertDate as cd
class WellfreshTmpl():
    if __name__ == "__main__":
        print("WELLFRESH CO.,LTD")

def getDataFromWellFresh(df, invoiceHeaderSheet, invoiceItem, i):
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
    indexIdx = -1
    articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    poNoIdx = -1
    amountIdx = -1
    # poNo = ''
    unit = ''
    # unitIdx = ''
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        info = str(excelValues[indexRow])
        if re.search(r'(?<=INVOICE NO:.)\w+', info) != None and isInvNoFound == False:
            isInvNoFound = True
            invNos = re.findall(r'(?<=INVOICE NO:.)\w+', info)
            invNo = "".join(invNos)
        if re.search(r'(?<=DATE:.)([\w]+)(\-)([\w]+)(\-)([\w]+)', info) != None and isIssueDateFound == False:
            isIssueDateFound = True
            issueDates = re.findall(r'(?<=DATE:.)([\w]+)(\-)([\w]+)(\-)([\w]+)', info)
            issueDate = "".join(issueDates[0])
        if re.search(r'Art NO.:', info) != None and re.search(r'QTY', info) != None and re.search(r'PRICE', info) != None and re.search(r'PO#', info) != None and isStartDescription == False:
            isStartDescription = True
            startDescriptionIdx = indexRow + 1
            indexCol = 0
            
            while indexCol < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][indexCol])
                # print(info)
                # print(re.search(r'TOTAL  AMOUNT \(USD\)', info))
                if info.strip() == 'NO.': 
                    indexIdx = indexCol
                if info == 'Art NO.:':
                    articleNoIdx = indexCol
                if info == 'FOB\nPRICE\n(USD)':
                    unitPriceIdx = indexCol
                if info == 'QTY\n(PCS)':
                    unit = 'PCS'
                    qtyIdx = indexCol
                if info == 'PO#':
                    poNoIdx = indexCol
                if re.search(r'TOTAL  AMOUNT \(USD\)', info) != None:
                    amountIdx = indexCol
                indexCol += 1
        if re.search(r'TOTAL', info) != None and isEndDescription == False:
            isEndDescription = True
            endDescriptionIdx = indexRow
        
        indexRow += 1
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    # print("Start : " + str(startDescriptionIdx) + " : " + str(endDescriptionIdx))
    while startDescriptionIdx < endDescriptionIdx:
        invoiceItemSheet.write('A' + str(worksheetIndex), str(excelValues[startDescriptionIdx][poNoIdx]))    
        invoiceItemSheet.write('B' + str(worksheetIndex), str(excelValues[startDescriptionIdx][indexIdx])) 
        invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
        invoiceItemSheet.write('D' + str(worksheetIndex), unit) 
        invoiceItemSheet.write('E' + str(worksheetIndex), str(excelValues[startDescriptionIdx][articleNoIdx])) 
        invoiceItemSheet.write('F' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitPriceIdx]))
        invoiceItemSheet.write('G' + str(worksheetIndex), str(round(excelValues[startDescriptionIdx][amountIdx], 2))) 
        worksheetIndex += 1
        startDescriptionIdx += 1
    total = str(round(excelValues[endDescriptionIdx][amountIdx], 2))
    grandTotal = total
    
    invoiceHeaderSheet.write('A' + sheetRow, customerNo)
    invoiceHeaderSheet.write('B' + sheetRow, invNo)
    invoiceHeaderSheet.write('C' + sheetRow, cd.getIssueDateWithoutMonthName(issueDate))
    invoiceHeaderSheet.write('D' + sheetRow, total)
    invoiceHeaderSheet.write('E' + sheetRow, extendTotal)
    invoiceHeaderSheet.write('F' + sheetRow, grandTotal)