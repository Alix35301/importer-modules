from getpass import getpass
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import getpass
import pandas as pd
import math
import pyodbc
import _datetime as datetime

username = getpass.getuser()
conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')


def getFile(sheetname):
    """
    Helper function to call gui function to extract file
    :return:
    dataframe

    """
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(parent=root)
    root.update()
    root.destroy()
    data = pd.read_excel(file_path, header=None, sheet_name=sheetname)
    return pd.DataFrame(data)


def check(l):
    """
    Function to check the string has values
    :param l:
    :return: string l
    """
    count = 1
    for i in l:
        if len(i) < 1 or i == 'nan':
            count += 1
    if count > len(l):
        return
    else:
        return l


def nullcheck(value):
    if str(value) == "NaN" or str(value) == "nan":
        return None
    elif float(value) == 0:
        return None
    else:
        return str(value)


def updateColumnData(c, colData, colDataOld, df, old_format_pd):
    colData.clear()
    [colData.append(data) for data in df.iloc[:160, c]]
    colDataOld.clear()
    [colDataOld.append(data) for data in old_format_pd.iloc[:218, c]]


checkValue = lambda v: None if v == 'nan' or v == '' else int(float(v))


def add_data_to_dataframe(datatodb):
    cls_id = []
    year = []
    instiution = []
    branches = []
    non_branches = []
    atm = []
    dpst_ins_holder = []
    dpst_acct_ins_plc = []
    borrower = []
    loan_acct = []
    cards = []
    outsd_ins_tech_res = []
    outsd_loan = []
    mobile_inet = []
    mobile_money = []
    created_by = []
    created_date = []

    for data in datatodb:
        for index, item in enumerate(data):
            if index == 0:
                cls_id.append(checkValue(item))
            elif index == 1:
                year.append(checkValue(item))
            elif index == 2:
                instiution.append(checkValue(item))
            elif index == 3:
                branches.append(checkValue(item))
            elif index == 4:
                non_branches.append(checkValue(item))
            elif index == 5:
                atm.append(checkValue(item))
            elif index == 6:
                dpst_ins_holder.append(checkValue(item))
            elif index == 7:
                dpst_acct_ins_plc.append(checkValue(item))
            elif index == 8:
                borrower.append(checkValue(item))
            elif index == 9:
                loan_acct.append(checkValue(item))
            elif index == 10:
                cards.append(checkValue(item))
            elif index == 11:
                outsd_ins_tech_res.append(checkValue(item))
            elif index == 12:
                outsd_loan.append(checkValue(item))
            elif index == 13:
                mobile_inet.append(checkValue(item))
            elif index == 14:
                mobile_money.append(checkValue(item))
            elif index == 15:
                created_by.append(username)
            elif index == 16:
                created_date.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    dbdf = pd.DataFrame({
        'cls_id': cls_id,
        'year': year,
        'institution': instiution,
        'branches': branches,
        'non_branches': non_branches,
        'atm': atm,
        'dpst_ins_holder': dpst_ins_holder,
        'dpst_acct_ins_plc': dpst_acct_ins_plc,
        'borrower': borrower,
        'loan_acct': loan_acct,
        'cards': cards,
        'outsd_ins_tech_res': outsd_ins_tech_res,
        'outsd_loan': outsd_loan,
        'mobile_inet': mobile_inet,
        'mobile_money': mobile_money,
        'created_by': created_by,
        'created_time': created_date
    })
    return dbdf


def makeConnection(dbname):
    conn_str = "Driver={SQL Server};Server=noo\\noo;Database=%s;Trusted_Connection=yes;" % dbname
    return pyodbc.connect(conn_str)


def truncate(dbname, table):
    conn = makeConnection(dbname)
    cursor = conn.cursor()
    print("Deleting old data..")
    tk_sql = "Delete from [{0}].[dbo].[{1}];DBCC CHECKIDENT ('[{0}].[dbo].[{1}]', RESEED, 0)".format(dbname, table)
    print(tk_sql)
    cursor.execute(tk_sql)
    conn.commit()


def get_classification(location, sheername):
    classification_xls = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FSI\classification.xlsx')
    file = pd.read_excel(classification_xls, sheet_name=sheername, header=None)
    return pd.DataFrame(file)


def checkDbf(dbname, table, date):
    conn = makeConnection(dbname)
    cursor = conn.cursor()
    sql = "SELECT * FROM [{0}].[dbo].[{1}]WHERE date > {2}".format(dbname, table, date)
    cursor.execute(sql)
    result = cursor.fetchall()
    if len(result) < 1:
        print("Data not found.. adding now!")
        return 1
    else:
        response = input("Data already exists! Would you like to overwrite the data ? (Y/N)")
        if response.upper() == "Y":
            return 2
            print("Overwriting data..")
        elif response.upper() == "N":
            print("ok")
            return 0


def remove_data(dbname, table, date):  # this removes yearly data
    conn = makeConnection(dbname)
    cursor = conn.cursor()
    sql = "DELETE FROM [{0}].[dbo].{1} WHERE date > '{2}' and date < '{3}'".format(dbname, table, date, int(date) + 1)
    print(sql)
    cursor.execute(sql)
    conn.commit()


def addToDb(dataframe, dbname, table):
    conn = makeConnection(dbname)
    cursor = conn.cursor()
    cols = ",".join([str(i) for i in dataframe.columns.tolist()])
    dataframe = dataframe.where(pd.notnull(dataframe), None)
    for i, row in dataframe.iterrows():
        sql = str("INSERT INTO [{0}].[dbo].{1} (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)").format(dbname,table)
        cursor.execute(sql, tuple(row))
        conn.commit()
    print("Done!", i, 'rows updated!!')
