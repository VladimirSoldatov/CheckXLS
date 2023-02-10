import pandas as pd
import openpyxl
import pyodbc
from datetime import datetime


class Sql:
    def __init__(self, database, server):
        self.cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                                   "Server=" + server + ";"
                                                        "Database=" + database + ";"
                                                                                 "Trusted_Connection=yes;")
        self.query = "-- {}\n\n-- Made in Python".format(datetime.now()
                                                         .strftime("%d/%m/%Y"))


# sql = Sql("webpbxReportDB", "I82Z0REPORT01")
# sql.cnxn()
# sql.query("select 1")
order = input("Enter order:")
file_name = input("Enter file name: ")
test_list = ['OptionsParamName', 'OptionsParamValue', 'OptionsCodeDefault',
             'TariffOptionsID', 'ValidityOfDiscount']

df = pd.read_excel("C:/CSV/исходник.xlsx", header=[0])
a = [j == j for i, j in zip(df.columns, df.columns) if i == j]
print(a)
ps = df.loc[:, "OptionsParamValue"]
print(df.head())
df = df.drop("OptionsParamValue", axis=1)
df.insert(2, "OptionsParamValue", ps)
df.insert(4, "Number", 2)
df.insert(6, "Date", "getdate()")
df.insert(7, "Order", "otrs " + order)
df.insert(8, "File_name", file_name)

df.to_excel("c:/csv/tmp.xlsx", index=False)
# for row in df.itertuples(name=None):
#    print(row)
