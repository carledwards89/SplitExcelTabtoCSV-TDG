import xlrd
import pandas as pd
import os
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox

sheetnames = []
issueNum = input("Enter the issue number: ")
def excel2csv(excel_file):
    # Open excel file
    workbook=xlrd.open_workbook(excel_file)
    # Get all sheet names
    sheet_names=workbook.sheet_names()
    for worksheet_name in sheet_names:
        sheetnames.append(worksheet_name)
        # Traverse each sheet and read it with Pandas

        data_xls=pd.read_excel(excel_file,worksheet_name,index_col=None)
        # Get the current directory of excel
        dir_path=os.path.abspath(os.path.dirname(excel_file))
        # Convert to csv and save to the csv folder in the directory where excel is located
        csv_path=dir_path+'\\csv\\'
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        data_xls.to_csv(csv_path+worksheet_name+'.csv',index=None,encoding='utf-8')

    for i in sheetnames:

        df = pd.read_csv("D:\python scripts\PEI\csv\\" +i+".csv")
        df.insert(5, "FullName", df["NAME 1"]+ " " + df["NAME 2"] + " " + df["NAME 3"])

        df["Length"] = ""
        df["Width"] = ""
        df["Height"] = ""
        df["TOTAL WEIGHT"] = df["TOTAL WEIGHT"] * 1000
        df["MAGAZINE"] = df["MAGAZINE"] + " " + issueNum
        df.loc[df["COUNTRY"] == "USA", "COUNTRY"] = "US"
        df.loc[df["COUNTRY"] == "Hong Kong", "COUNTRY"] = "HK"
        df.loc[df["COUNTRY"] == "Singapore", "COUNTRY"] = "SG"
        df.loc[df["COUNTRY"] == "China", "COUNTRY"] = "CN"
        df.loc[df["COUNTRY"] == "Luxembourg", "COUNTRY"] = "LU"

        df.to_csv("D:\python scripts\PEI\csv\\" + i+".csv", index=False)


excel2csv(r'D:\python scripts\PEI\PEI.xlsx')

app = QtWidgets.QApplication([])

dlg = uic.loadUi("CookAPP.ui")

dlg.createCSV.clicked.connect(excel2csv)


dlg.show()
app.exec()