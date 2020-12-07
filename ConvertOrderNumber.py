import re
class ConvertOrderNumber():
    if __name__ == "__main__":
        print("ConvertOrderNumber")

def getOrderNumberValue(orderNumber):
    
    
    if re.search(r'(?<=PO_IMP/)\d+', orderNumber) != None:
        orderNumber = re.findall(r'(?<=PO_IMP/)\d+', orderNumber)
        print(orderNumber[0].strip())
        return orderNumber[0].strip()
    return orderNumber