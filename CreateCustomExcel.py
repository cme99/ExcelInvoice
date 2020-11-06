import xlsxwriter
class FoshanTiansiTmpl():
    if __name__ == "__main__":
        print("FoshanTiansi")
def initSheet(invoiceItem, sheet_name):
    invoiceItemSheet = invoiceItem.add_worksheet(sheet_name)
    invoiceItemSheet.write('A1', 'OrderNumber')
    invoiceItemSheet.write('B1', 'Index')
    invoiceItemSheet.write('C1', 'Quantity')
    invoiceItemSheet.write('D1', 'Unit')
    invoiceItemSheet.write('E1', 'ArticleNo')
    invoiceItemSheet.write('F1', 'Price')
    invoiceItemSheet.write('G1', 'Amount')
    invoiceItemSheet.write('H1', 'Extended')
    return (invoiceItem, invoiceItemSheet)