import pandas as pd
import xlsxwriter
import os
import sys
import numpy as np
import re
import ZhongShanGuanYiImportExportTradeTmpl as zsgyTmpl
import FuranTradingTmpl as ftTmpl
import ZhongShanGuanQinTmpl as zsgqTmpl
import FoshanTiansiTmpl as fotTmpl
import IMagicKitchenAppliancesTmpl as imkaTmpl
import WellseeEnterpriseTmpl as wsepTmpl
import WellfreshTmpl as wfTmpl

types = ['FURUAN TRADING CO.,LTD.OF KAIPING CITY', 
'ZHONGSHAN GUANGQIN TRADE CO.,LTD.', 
'Foshan Tiansi Hardware Co.,Ltd',
'ZHONGSHAN GUANGYI IMPORT AND EXPORT TRADE CO., LTD.',
'iMagic Kitchen Appliances(H.K.) Co., Ltd.',
'WELLSEE ENTERPRISE CO., LTD.',
'WELLFRESH CO.,LTD'
]

worksheetIndex = 2
def getExcelTemplate(columns, types):
    for column in columns:
        for type in types:
            if (column.strip() == type):
                return column.strip()
    return ''

def checkType():
    for row in df.values:
        for type in types:
            if re.search(type, str(row)) != None:
               return str(type)
    return ''

def goToFuncBasedOnExcelType(excelType, i):
    
    if excelType.upper() == 'FURUAN TRADING CO.,LTD.OF KAIPING CITY'.upper():
        print ("FURUAN TRADING CO.,LTD.OF KAIPING CITY Template")
        ftTmpl.getDataFromFuranTrading(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'ZHONGSHAN GUANGQIN TRADE CO.,LTD.'.upper():
        print ("ZHONGSHAN GUANGQIN TRADE CO.,LTD. Template")
        zsgqTmpl.getDataFromZhongShanGuanQin(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'Foshan Tiansi Hardware Co.,Ltd'.upper():
        print ("Foshan Tiansi Hardware Co.,Ltd Template")
        fotTmpl.getDataFromFoshanTiansi(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'ZHONGSHAN GUANGYI IMPORT AND EXPORT TRADE CO., LTD.'.upper():
        print ("ZHONGSHAN GUANGYI IMPORT AND EXPORT TRADE CO., LTD. Template")
        zsgyTmpl.getDataFromZhongShanGuanYiImportExportTrade(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'iMagic Kitchen Appliances(H.K.) Co., Ltd.'.upper():
        print ("iMagic Kitchen Appliances(H.K.) Co., Ltd. Template")
        imkaTmpl.getDataFromIMagicKitchenAppliances(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'WELLSEE ENTERPRISE CO., LTD.'.upper():
        print("WELLSEE ENTERPRISE CO., LTD.")
        wsepTmpl.getDataFromWellseeEnterprise(df, invoiceHeaderSheet, invoiceItem, i)
    elif excelType.upper() == 'WELLFRESH CO.,LTD'.upper():
        print("WELLFRESH CO.,LTD")
        wfTmpl.getDataFromWellFresh(df, invoiceHeaderSheet, invoiceItem, i)
    else: 
        excelType = checkType()
        goToFuncBasedOnExcelType(excelType, i)
        

def initWorkBook():
    invoiceHeader = xlsxwriter.Workbook("InvoiceHeader.xlsx")  
    invoiceHeaderSheet = invoiceHeader.add_worksheet("Header")
    invoiceHeaderSheet.write("A1", "Customer No")
    invoiceHeaderSheet.write("B1", "Invoice No")
    invoiceHeaderSheet.write("C1", "Date")
    invoiceHeaderSheet.write("D1", "Total")
    invoiceHeaderSheet.write("E1", "Extend Total")
    invoiceHeaderSheet.write("F1", "Grand Total")
    # invoiceHeaderSheet.write("G1", "File name")
    invoiceItem = xlsxwriter.Workbook("InvoiceItem.xlsx")
    return (invoiceHeader, invoiceHeaderSheet, invoiceItem)

if __name__ == "__main__":
    entries = os.listdir(sys.argv[1])
    i = 2
    (invoiceHeader, invoiceHeaderSheet, invoiceItem) = initWorkBook()

    for entry in entries:
        if entry.lower().endswith('.xls') or entry.lower().endswith('.xlsx'):
            print(os.getcwd() + "\\" + sys.argv[1] + "\\" + entry)
            file_name = os.getcwd() + "\\" + sys.argv[1] + "\\" + entry
            # cropPDF(file_name, 1, 1, 100, 100)
            # path = r''+strPath
            df = pd.read_excel(file_name) 
            excelType = getExcelTemplate(df.columns, types)
            goToFuncBasedOnExcelType(excelType, str(i))
            # invoiceHeaderSheet.write("G" + str(i), file_name)
            i+=1
    # strPath = sys.argv[1]
    invoiceItem.close()
    invoiceHeader.close()
    
