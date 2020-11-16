import re
import datetime
import ConvertDate as cd
import CreateCustomExcel as ccx
class WellseeEnterpriseTmpl():
    if __name__ == "__main__":
        print("WELLSEE ENTERPRISE CO., LTD.")

multiChars = "'[]"
def removeNoise(srcString):
    for character in multiChars:
        srcString = srcString.replace(character, "")
    return srcString.replace('nan', '')

def getDataFromWellseeEnterprise(df, invoiceHeaderSheet, invoiceItem, i):
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
    # articleNoIdx = -1
    qtyIdx = -1
    unitPriceIdx = -1
    poNoIdx = -1
    amountIdx = -1
    poNo = ''
    # unit = ''
    unitIdx = -1
    worksheetIndex = 2
    excelValues = df.values
    while indexRow < len(excelValues):
        info = str(excelValues[indexRow])
        if re.search(r'Date:', info) != None and isIssueDateFound == False:
            # print(info)
            isIssueDateFound = True
            i = 0
            while i < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][i])
                if re.search(r'Date', info):
                    issueDates = str(excelValues[indexRow][i+1])
                    issueDate = removeNoise(''.join(issueDates))
                    # print("Issue Date : " + issueDate)
                    issueDate = str(issueDate).split(' ')[0]
                    break
                i += 1
        if re.search(r'Ref.No.', info) != None and isInvNoFound == False:
            # print(info)
            isInvNoFound = True
            i = 0
            while i < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][i])
                if re.search(r'Ref.No.', info):
                    invNo = str(excelValues[indexRow][i+1])
                    break
                i += 1
        if re.search(r'ITEM#', info) != None and re.search(r'Q\'TY', info) != None and re.search(r'UNIT', info) != None and re.search(r'PRICE US\$', info) != None and re.search(r'AMOUNT US\$', info) != None and isStartDescription == False:
            isStartDescription = True
            startDescriptionIdx = indexRow + 1
            while indexCol < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][indexCol])
                if re.search(r'ITEM#', info) != None: 
                    poNoIdx = indexCol
                    # articleNoIdx = indexCol
                if re.search(r'Q\'TY', info) != None: 
                    qtyIdx = indexCol
                if re.search(r'UNIT', info) != None:
                    unitIdx = indexCol
                if info == 'PRICE US$':
                    unitPriceIdx = indexCol
                if info == 'AMOUNT US$':
                    amountIdx = indexCol
                indexCol += 1
            
        if re.search(r'(?<=SAY TOTAL US DOLLARS)', info) != None and isEndDescription == False:
            isEndDescription = True
            endDescriptionIdx = indexRow - 1

        indexRow += 1
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    while startDescriptionIdx < endDescriptionIdx:
        poStr = str(excelValues[startDescriptionIdx][poNoIdx])
        
        if re.search(r'(?<=PO_IMP/|PO_SER/).\w+', poStr) != None:
            info = re.findall(r'(?<=PO_IMP/|PO_SER/).\w+', poStr)
            poNo = ', '.join(info).strip()
        elif re.search(r'REPLACEMENT', poStr) !=None:
            poNo = 'REPLACEMENT'
        else:  
            # print(poStr)
            if re.search(r'[0-9]+\.[0-9]+\.[0-9]+', poStr) != None:
                # print(re.search(r'[0-9]+\.[0-9]+\.[0-9]+', poStr))
                articleNo = ''.join(re.findall(r'[0-9]+\.[0-9]+\.[0-9]+', poStr)).strip()
            else:
                articleNo = poStr.split(' ')[0].strip()

            invoiceItemSheet.write('A' + str(worksheetIndex), poNo)    
            invoiceItemSheet.write('B' + str(worksheetIndex), str(worksheetIndex - 1)) 
            invoiceItemSheet.write('C' + str(worksheetIndex), str(excelValues[startDescriptionIdx][qtyIdx]))
            invoiceItemSheet.write('D' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitIdx])) 
            invoiceItemSheet.write('E' + str(worksheetIndex), articleNo) 
            invoiceItemSheet.write('F' + str(worksheetIndex), str(excelValues[startDescriptionIdx][unitPriceIdx]))
            invoiceItemSheet.write('G' + str(worksheetIndex), str(excelValues[startDescriptionIdx][amountIdx])) 
            worksheetIndex += 1
        startDescriptionIdx += 1
    total = str(round(excelValues[endDescriptionIdx][amountIdx], 2))
    grandTotal = total
    issueDate = issueDate.strip()
    invoiceHeaderSheet.write('A' + sheetRow, customerNo.strip())
    invoiceHeaderSheet.write('B' + sheetRow, invNo.strip())
    invoiceHeaderSheet.write('C' + sheetRow, cd.getIssueDateWithoutMonthName(issueDate))
    invoiceHeaderSheet.write('D' + sheetRow, total)
    invoiceHeaderSheet.write('E' + sheetRow, extendTotal)
    invoiceHeaderSheet.write('F' + sheetRow, grandTotal)
        