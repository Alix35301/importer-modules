from getpass import getpass
from tkinter import filedialog
import tkinter as tk

import MySQLdb
import pandas as pd
import getpass
from datetime import datetime

def connectToDB():
    # conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=LOANS;Trusted_Connection=yes;')
    conn = MySQLdb.connect (host = "localhost",
                        user = "root",
                        passwd = "1234",
                        db = "tourism")
    return conn

import pyodbc
#
# username = getpass.getuser()
# conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
#
#
# def getFile():
#     """
#     Helper function to call gui function to extract file
#     :return:
#     file_name
#
#     """
#     root = tk.Tk()
#     root.withdraw()
#     root.attributes('-topmost', 1)
#     file_path = filedialog.askopenfilename(parent=root)
#     root.update()
#     root.destroy()
#     return file_path
#
#
# def check(l):
#     """
#     Function to check the string has values
#     :param l:
#     :return: string l
#     """
#     count = 1
#     for i in l:
#         if len(i) < 1 or i == 'nan':
#             count += 1
#     if count > len(l):
#         return
#     else:
#         return l
#
#
# def updateColumnData(c, colData, colDataOld, df, old_format_pd):
#     colData.clear()
#     [colData.append(data) for data in df.iloc[:160, c]]
#     colDataOld.clear()
#     [colDataOld.append(data) for data in old_format_pd.iloc[:218, c]]
#
#
# checkValue = lambda v: None if v == 'nan' or v == '' else int(float(v))
#
#
# def add_data_to_dataframe(datatodb):
#     cls_id = []
#     year = []
#     instiution = []
#     branches = []
#     non_branches = []
#     atm = []
#     dpst_ins_holder = []
#     dpst_acct_ins_plc = []
#     borrower = []
#     loan_acct = []
#     cards = []
#     outsd_ins_tech_res = []
#     outsd_loan = []
#     mobile_inet = []
#     mobile_money = []
#     created_by = []
#     created_date = []
#
#     for data in datatodb:
#         for index, item in enumerate(data):
#             if index == 0:
#                 cls_id.append(checkValue(item))
#             elif index == 1:
#                 year.append(checkValue(item))
#             elif index == 2:
#                 instiution.append(checkValue(item))
#             elif index == 3:
#                 branches.append(checkValue(item))
#             elif index == 4:
#                 non_branches.append(checkValue(item))
#             elif index == 5:
#                 atm.append(checkValue(item))
#             elif index == 6:
#                 dpst_ins_holder.append(checkValue(item))
#             elif index == 7:
#                 dpst_acct_ins_plc.append(checkValue(item))
#             elif index == 8:
#                 borrower.append(checkValue(item))
#             elif index == 9:
#                 loan_acct.append(checkValue(item))
#             elif index == 10:
#                 cards.append(checkValue(item))
#             elif index == 11:
#                 outsd_ins_tech_res.append(checkValue(item))
#             elif index == 12:
#                 outsd_loan.append(checkValue(item))
#             elif index == 13:
#                 mobile_inet.append(checkValue(item))
#             elif index == 14:
#                 mobile_money.append(checkValue(item))
#             elif index == 15:
#                 created_by.append(username)
#             elif index == 16:
#                 created_date.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
#     dbdf = pd.DataFrame({
#         'cls_id': cls_id,
#         'year': year,
#         'institution': instiution,
#         'branches': branches,
#         'non_branches': non_branches,
#         'atm': atm,
#         'dpst_ins_holder': dpst_ins_holder,
#         'dpst_acct_ins_plc': dpst_acct_ins_plc,
#         'borrower': borrower,
#         'loan_acct': loan_acct,
#         'cards': cards,
#         'outsd_ins_tech_res': outsd_ins_tech_res,
#         'outsd_loan': outsd_loan,
#         'mobile_inet': mobile_inet,
#         'mobile_money': mobile_money,
#         'created_by': created_by,
#         'created_time': created_date
#     })
#     return dbdf
#
#
# def truncate(db,table):
#     conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
#     cursor = conn.cursor()
#     print("Deleting old data..")
#     sql =("Delete"
#           "from {};")
#     sql1 = "Delete from [{0}].[dbo].[{1}];".format(db,table)
#     sql2 ="DBCC CHECKINDENT('[{0}].[dbo].[{1}]',RESEED,0)".format(db,table)
#     sql = sql1+sql2
#     # tk_sql = ("""Delete from [FAS].[dbo].[DATA];DBCC CHECKIDENT ('[FAS].[dbo].[DATA]', RESEED, 0)""")
#     cursor.execute(sql)
#     conn.commit()
#     print("done")
#
#
# truncate("FAS","DATA")
def addToDb(dataframe, dbname, table):
    conn = connectToDB()
    cursor = conn.cursor()
    cols = ",".join([str(i) for i in dataframe.columns.tolist()])
    dataframe = dataframe.where(pd.notnull(dataframe), None)
    for i, row in dataframe.iterrows():
        sql = str("INSERT INTO {0}.{1} (" + cols + ") VALUES (" + "%s," * (len(row) - 1) + "%s)").format(dbname,table)
        # print(row)
        cursor.execute(sql, tuple(row))
        conn.commit()
    # print("Done!", i, 'rows updated!!')
