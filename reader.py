import pandas as pd
import xlsxwriter
import os
import sys
import numpy as np
import re
import shutil
import PDFScanTmpl as pdfScan
import ZhongShanGuanYiImportExportTradeTmpl as zsgyTmpl
import FuranTradingTmpl as ftTmpl
import ZhongShanGuanQinTmpl as zsgqTmpl
import FoshanTiansiTmpl as fotTmpl
import IMagicKitchenAppliancesTmpl as imkaTmpl
import WellseeEnterpriseTmpl as wsepTmpl
import WellfreshTmpl as wfTmpl
import subprocess

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

def checkType(df):
    for row in df.values:
        for type in types:
            if re.search(type, str(row)) != None:
               return str(type)
    return ''

def goToFuncBasedOnExcelType(excelType, invoiceHeaderSheet, invoiceItem, df, i):
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
        excelType = checkType(df)
        goToFuncBasedOnExcelType(excelType, invoiceHeaderSheet, invoiceItem, df, i)

def initWorkBook(folder_path):
    headerExcelFile = "InvoiceHeader.xlsx"
    invoiceHeader = xlsxwriter.Workbook(headerExcelFile)  
    headerRow = 1
    invoiceHeaderSheet = invoiceHeader.add_worksheet("Header")
    ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    headerDatas = pd.DataFrame({})
    if os.path.exists(headerExcelFile):
        # invoiceHeaderSheet = invoiceHeader.get_worksheet_by_name("Header")
        headerDatas = pd.read_excel(headerExcelFile)  
        print(headerDatas.values)
        headerRow = len(headerDatas.values)
        index = 0
        # # print(headerDatas.columns)
        for h in headerDatas.columns:
            print(ascii_uppercase[index]+str(1))
            invoiceHeaderSheet.write(ascii_uppercase[index]+str(1), h.strip())
            index+=1
        indexRow = 0
        
        while indexRow < len(headerDatas.values):
            print(headerDatas.values[indexRow])
            indexCol = 0
            while indexCol < len(headerDatas.values[indexRow]):
                info = str(headerDatas.values[indexRow][indexCol])
                if info == 'nan':
                    info = ''
                invoiceHeaderSheet.write(ascii_uppercase[indexCol]+str(indexRow+2), info)
                indexCol += 1
            indexRow+=1
        # headerRow += 1
        # invoiceItem = xlsxwriter.Workbook(folder_path + "\\Output\\InvoiceItem.xlsx")
        # return (invoiceHeader, invoiceHeaderSheet, invoiceItem, headerRow)
    else:
        invoiceHeaderSheet.write("A"+str(headerRow), "Customer No")
        invoiceHeaderSheet.write("B"+str(headerRow), "Invoice No")
        invoiceHeaderSheet.write("C"+str(headerRow), "Date")
        invoiceHeaderSheet.write("D"+str(headerRow), "Total")
        invoiceHeaderSheet.write("E"+str(headerRow), "Extend Total")
        invoiceHeaderSheet.write("F"+str(headerRow), "Grand Total")
        headerRow += 1
        # invoiceHeaderSheet.write("G1", "File name")
    invoiceItem = xlsxwriter.Workbook("InvoiceItem.xlsx")
    return (invoiceHeader, invoiceHeaderSheet, invoiceItem, headerDatas, headerRow)

def execute(folder_path, excel_path):
    (invoiceHeader, invoiceHeaderSheet, invoiceItem, headerDatas, headerRow) = initWorkBook(folder_path)
    
    entries = os.listdir(os.getcwd() + "\\" +folder_path +"\\"+ excel_path)
    print(str(headerRow))
    i = headerRow
    for entry in entries:
        if entry.lower().endswith('.xls') or entry.lower().endswith('.xlsx'):            
            file_name = os.getcwd() + "\\" +folder_path +"\\"+ excel_path +"\\" + entry
            print(file_name)
            df = pd.read_excel(file_name) 
            excelType = getExcelTemplate(df.columns, types)
            goToFuncBasedOnExcelType(excelType, invoiceHeaderSheet, invoiceItem, df, str(i))
            i+=1
    invoiceItem.close()
    invoiceHeader.close()
    return "Execute"

if __name__ == "__main__":
    execute(sys.argv[1], "")
