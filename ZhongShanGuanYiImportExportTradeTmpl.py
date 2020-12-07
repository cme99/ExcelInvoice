import re
import CreateCustomExcel as ccx
import ConvertDate as cd
class ZhongShanGuanYiImportExportTradeTmpl():
    if __name__ == "__main__":
        print("Zhong Shan Guan Yi Import and Export Trade")

multiChars = "'[]]nan"
def removeNoise(srcString):
    for character in multiChars:
        srcString = srcString.replace(character, "")
    return srcString

def getDataFromZhongShanGuanYiImportExportTrade(df, invoiceHeaderSheet, invoiceItem, i):
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
    # poNoIdx = -1
    amountIdx = -1
    poNo = ''
    unit = ''
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        # infoFirstCol = str(excelValues[indexRow][0])
        info = str(excelValues[indexRow])
        if re.search(r'(?<=FUYU Item)', info) != None and re.search(r'(?<=HAFELE Item)', info) != None and re.search(r'(?<=Unit Price)', info) != None and re.search(r'(?<=Quantity)', info) != None and isStartDescription == False:
            isStartDescription = True
            startDescriptionIdx = indexRow + 1
            while indexCol < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][indexCol])
                if info == 'FUYU Item': 
                    poIdx = indexCol
                if info == 'HAFELE Item':
                    articleNoIdx = indexCol
                if info == 'Unit Price':
                    unitPriceIdx = indexCol
                if info == 'Quantity':
                    qtyIdx = indexCol
                if info == 'Amount':
                    amountIdx = indexCol
                indexCol += 1
        if re.search(r'(?<=Total:)', info) != None and isEndDescription == False:
            isEndDescription = True
            endDescriptionIdx = indexRow

        if re.search(r'(?<=Invoice No :\s).*(?=\')', info) != None and isInvNoFound == False:
            isInvNoFound = True
            invNos = re.findall(r'(?<=Invoice No :\s).*(?=\')', info)
            invNo = ''.join(invNos)
            # invoiceNoIdx = indexRow
            
        if re.search(r'(?<=Date :\s).*(?=\')', info) != None and isIssueDateFound == False:
            print("Date : " + info)
            isIssueDateFound = True
            issueDates = re.findall(r'(?<=Date :\s).*(?=\')', info)[0]
            issueDate = ''.join(issueDates)
            # issueDateIdx = indexRow

        if isInvNoFound and isIssueDateFound:
            pass   
        indexRow += 1
    poNo = ''
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    while startDescriptionIdx < endDescriptionIdx:
        poStr = str(excelValues[startDescriptionIdx][poIdx])
        
        if re.search(r'(?<=PO No.:\s)\d+', poStr) != None:
            info = re.findall(r'(?<=PO No.:\s)\d+', poStr)
            poNo = ''.join(info).strip()
        else: 
            invoiceItemSheet.write('A' + str(worksheetIndex), poNo)    
            invoiceItemSheet.write('B' + str(worksheetIndex), str(worksheetIndex - 1)) 
            invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
            invoiceItemSheet.write('D' + str(worksheetIndex), unit) 
            invoiceItemSheet.write('E' + str(worksheetIndex), str(excelValues[startDescriptionIdx][articleNoIdx])) 
            invoiceItemSheet.write('F' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitPriceIdx]))
            invoiceItemSheet.write('G' + str(worksheetIndex), str(round(excelValues[startDescriptionIdx][amountIdx], 2))) 
            worksheetIndex += 1
        startDescriptionIdx += 1
    total = str(round(excelValues[endDescriptionIdx][amountIdx], 2))
    grandTotal = total
    invoiceHeaderSheet.write('A' + sheetRow, customerNo.strip())
    invoiceHeaderSheet.write('B' + sheetRow, invNo.strip())
    invoiceHeaderSheet.write('C' + sheetRow, cd.getIssueDate(issueDate.strip()))
    invoiceHeaderSheet.write('D' + sheetRow, total)
    invoiceHeaderSheet.write('E' + sheetRow, extendTotal)
    invoiceHeaderSheet.write('F' + sheetRow, grandTotal)
        