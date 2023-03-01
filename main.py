from datetime import datetime
import pandas as pd
import pyodbc
from tkinter import Tk
from tkinter.filedialog import askopenfilename


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
Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
file_name = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
print(file_name)

pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 10)
# files = []
# path = "C:/CSV/"
# os.chdir(path)
order = input("Enter order:")
# for index, file in enumerate(glob.glob("*.xlsx")):
#    print(index, " ", file)
#    files.append(file)
# file_index = input("Enter file name index: ")
# file_name = files[int(file_index)]
test_list = ['OptionsParamName', 'OptionsParamValue', 'OptionsCodeDefault',
             'TariffOptionsID', 'ValidityOfDiscount']
df = pd.read_excel(file_name, header=[0])
# df = pd.read_excel("C:/CSV/source.xlsx", header=[0])
flag = True
for i, j in zip(df.columns, test_list):
    if i != j:
        flag = False
if flag:
    print(df.head())
    cols = list(df.columns)
    a, b, c, d, e = cols.index('OptionsParamName'), cols.index('OptionsParamValue'), cols.index(
        'TariffOptionsID'), cols.index('OptionsCodeDefault'), cols.index('ValidityOfDiscount')
    cols[a], cols[b], cols[c], cols[d], cols[e] = cols[a], cols[b], cols[d], cols[c], cols[e]
    df = df[cols]
    df.columns = ['OptionsParamName', 'OptionsParamValue', 'TariffOptionsID',
                  'OptionsCode', 'ValidityOfDiscount']
    df["OptionsParamName"] = df["OptionsParamName"].str.replace("sn", "'sn'")
    df["OptionsParamValue"] = [f"'{i}'" for i in df["OptionsParamValue"]]
    df.fillna('NULL', inplace=True)
    print(df.head())
    df["Result"] = 0
    df["recordDateAdd"] = "getdate()"
    df["description_inc"] = "'otrs " + order + "'"
    df["description_act"] = "'" + file_name.split('/')[-1] + "'"
    print(df.head())
    df.to_csv("c:/csv/tmp.csv", index=False, header=None, encoding='utf-8')
    text = open("c:/csv/tmp.csv")
    while text.read():
        print(text.readline())
else:
    print("Unknown file")

# for row in df.itertuples(name=None):
#    print(row)
