import _datetime
from pathlib import Path
from tkinter import filedialog
import getpass
import math
import numpy as np
import os
import pandas as pd
import pyodbc
import datetime
import tkinter as tk

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(parent=root)
    root.update()
    root.destroy()

    file_name = Path(file_path).name
    fc = file_name.split("_")[0]
    date = datetime.datetime.today()
    date = date.strftime("%Y-%m-%d")
    username = getpass.getuser()




    data = pd.read_excel(file_path, header=None, sheet_name="Annex 3") #changed
    df = pd.DataFrame(data)

    classification_xls = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FSI\classification.xlsx')
    cls = pd.read_excel(classification_xls, sheet_name="DEPOSIT", header=None)  #changed
    df_cls = pd.DataFrame(cls)

    codes_from_classification = df_cls.iloc[:, 3].to_list()

    cls_ids=[]
    years = []
    q1_values =[]
    q2_values =[]
    q3_values =[]
    q4_values =[]
    created_time=[]
    created_by=[]
    data_string =[]

    def nullcheck(value):
        if str(value) =="NaN" or str(value) =="nan":
            return None
        elif float(value) ==0:
            return None
        else:
            return str(value)
    scale =1000000
    row = 10
    q1_value = ""
    q2_value = ""
    q3_value = ""
    q4_value = ""
    for code in df.iloc[:, 2]:
        if str(code) != "nan" and code in codes_from_classification:
            if code in codes_from_classification:
                data_to_list = df.iloc[9, 6:].values.tolist()
                value_list = df.iloc[row, 6:].values.tolist()
                # iterating year lable year eg: 2015Q1
                for data in data_to_list:
                    if str(data) != "nan" and str(data) != "NaN":
                        year = data[:4]
                        qt = data[4:]
                        if qt == "Q1":
                            q1_value =value_list[data_to_list.index(str(data))]
                        elif qt =="Q2":
                            q2_value =value_list[data_to_list.index(str(data))]
                        elif qt =="Q3":
                            q3_value =value_list[data_to_list.index(str(data))]
                        elif qt =="Q4":
                            q4_value =value_list[data_to_list.index(str(data))]

                            temp = str(codes_from_classification.index(str(code))) +\
                                   ","+ str(year)+\
                                   ","+str(nullcheck(float(q1_value)*scale))+\
                                   ","+str(nullcheck(float(q2_value)*scale))+\
                                   ","+str(nullcheck(float(q3_value)*scale))+\
                                   ","+str(nullcheck(float(q4_value)*scale))
                            data_string.append(temp)


                row = row + 1


    for i in data_string:
        l = i.split(",")
        if(l[2]!="None" and l[3]!="None" and l[4]!="None" and l[5]!="None"):
            cls_ids.append(l[0])
            years.append(l[1])
            created_time.append(_datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            created_by.append(username)
            if l[2] !="None":
                q1_values.append(l[2])
            else:
                q1_values.append(None)

            if l[3] !="None":
                q2_values.append(l[3])
            else:
                q2_values.append(None)
            if l[4] !="None":
                q3_values.append(l[4])
            else:
                q3_values.append(None)
            if l[5] !="None":
                q4_values.append(l[5])
            else:
                q4_values.append(None)


    dbdf = pd.DataFrame({
        'cls_id': cls_ids,
        'year': years,
        'q1': q1_values,
        'q2': q2_values,
        'q3': q3_values,
        'q4': q4_values,
        'created_by': created_by,
        'created_time': created_time
    })
    #
    # conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FSI;Trusted_Connection=yes;')
    #
    # cursor = conn.cursor()
    # print("Deleting old data..")
    # tk_sql = "Delete from [FSI].[dbo].[DATA];DBCC CHECKIDENT ('[FSI].[dbo].[data]', RESEED, 0)"
    # cursor.execute(tk_sql)
    # cols = ",".join([str(i) for i in dbdf.columns.tolist()])
    #
    # dbdf = dbdf.where(pd.notnull(dbdf), None)
    # for i, row in dbdf.iterrows():
    #     sql = "INSERT INTO [FSI].[dbo].DATA (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)"
    #     cursor.execute(sql, tuple(row))
    #     conn.commit()
    # print("Done!", i, 'rows updated!!')