import re
import datetime
import ConvertDate as cd
import CreateCustomExcel as ccx
class IMagicKitchenAppliancesTmpl():
    if __name__ == "__main__":
        print("iMagic Kitchen Appliances(H.K.) Co., Ltd.")

multiChars = "'[]"
def removeNoise(srcString):
    for character in multiChars:
        srcString = srcString.replace(character, "")
    return srcString.replace('nan', '')

def allNan(stringList):
    for string in stringList:
        if string != 'nan':
            return False
    return True

def isfloat(str):
    return isinstance(str, float)

def getDataFromIMagicKitchenAppliances(df, invoiceHeaderSheet, invoiceItem, i):
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
        
        value = str(excelValues[indexRow])
       
        info = removeNoise(''.join(value)) 
        
        if re.search(r'(?<=INVOICE NO.:).*', info) != None and isInvNoFound == False:
            isInvNoFound = True
            invNos = re.findall(r'(?<=INVOICE NO.:).*', info)
            invNo = ''.join(invNos)
            # invoiceNoIdx = indexRow
        if re.search(r'(?<=DATE:).*', info) != None and isIssueDateFound == False:
            # print(info)
            isIssueDateFound = True
            issueDates = re.findall(r'(?<=DATE:).*', info)
            issueDate = removeNoise(''.join(issueDates))
            if re.search(r'(?<=datetime.datetime).*', issueDate) != None:
                issueDates = re.findall(r'\((.*)\)', issueDate)
                issueDate = ''.join(issueDates).split(',')
                issueDate = datetime.datetime(int(issueDate[0]), int(issueDate[1]), int(issueDate[2]))
                issueDate = str(issueDate).split(' ')[0]
            # issueDateIdx = indexRow
        
        if re.search(r'(?<=Item)', info) != None and re.search(r'(?<=Article No.)', info) != None and re.search(r'(?<=U\/Price\\n\(USD\))', info) != None and re.search(r'(?<=Qty\\n\(PC\))', info) != None and isStartDescription == False:
            isStartDescription = True
            startDescriptionIdx = indexRow + 1
            while indexCol < len(excelValues[indexRow]):
                info = str(excelValues[indexRow][indexCol])
                if re.search(r'Item', info) != None: 
                    poIdx = indexCol
                if info == 'Article No.':
                    articleNoIdx = indexCol
                if info == 'U/Price\n(USD)':
                    unitPriceIdx = indexCol
                if info == 'Qty\n(PC)':
                    unit = 'PC'
                    qtyIdx = indexCol
                if info == 'Total Amount\n(USD)':
                    amountIdx = indexCol
                indexCol += 1
        if re.search(r'(?<=SAY US DOLLARS)', info) != None and isEndDescription == False:
            isEndDescription = True
            endDescriptionIdx = indexRow - 1

        indexRow += 1
    (invoiceItem, invoiceItemSheet) = ccx.initSheet(invoiceItem, invNo.strip())
    poNo = ''
    while startDescriptionIdx < endDescriptionIdx:
        poStr = str(excelValues[startDescriptionIdx][poIdx])
        # print(poStr)
        if re.search(r'(?<=PO_IMP/).\w+', poStr) != None:
            info = re.findall(r'(?<=PO_IMP/).\w+', poStr)
            poNo = ', '.join(info).strip()
            index = poNo
        else: 
            # print([str(excelValues[startDescriptionIdx][articleNoIdx]), str(excelValues[startDescriptionIdx][qtyIdx]), str(excelValues[startDescriptionIdx][unitPriceIdx])])
            if not allNan([str(excelValues[startDescriptionIdx][articleNoIdx]), str(excelValues[startDescriptionIdx][qtyIdx]), str(excelValues[startDescriptionIdx][unitPriceIdx])]):
                articleNo = articleNo if (str(excelValues[startDescriptionIdx][articleNoIdx]) == 'nan') else str(excelValues[startDescriptionIdx][articleNoIdx])
                if re.search(r'[0-9]+\.[0-9]+\.[0-9]+', articleNo) != None:
                    articleNo = ''.join(re.findall(r'[0-9]+\.[0-9]+\.[0-9]+', articleNo)).strip()
                qty = qty if (str(excelValues[startDescriptionIdx][qtyIdx]) == 'nan') else str(excelValues[startDescriptionIdx][qtyIdx])
                unitPrice = unitPrice if (str(excelValues[startDescriptionIdx][unitPriceIdx]) == 'nan') else str(excelValues[startDescriptionIdx][unitPriceIdx])
                # invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
                if str(excelValues[startDescriptionIdx][poIdx]) == 'nan':
                    invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
                else:
                    index = str(excelValues[startDescriptionIdx][poIdx])
                    invoiceItemSheet.write('A' + str(worksheetIndex), poNo)
                if str(excelValues[startDescriptionIdx][poIdx]) == 'nan':
                    invoiceItemSheet.write('B' + str(worksheetIndex), index)
                else:
                    index = str(excelValues[startDescriptionIdx][poIdx])
                    invoiceItemSheet.write('B' + str(worksheetIndex), index)
                invoiceItemSheet.write('C' + str(worksheetIndex), qty)
                invoiceItemSheet.write('D' + str(worksheetIndex), unit)
                invoiceItemSheet.write('E' + str(worksheetIndex), articleNo) 
                invoiceItemSheet.write('F' + str(worksheetIndex), unitPrice)
                amount = excelValues[startDescriptionIdx][amountIdx]
                if isfloat(amount):
                    amount = round(excelValues[startDescriptionIdx][amountIdx], 2)
                    # print(str(amount))
                    invoiceItemSheet.write('G' + str(worksheetIndex), str(amount)) 
                else:
                    invoiceItemSheet.write('G' + str(worksheetIndex), str(excelValues[startDescriptionIdx][amountIdx]))
                
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