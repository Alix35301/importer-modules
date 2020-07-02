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

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', 1)
file_path = filedialog.askopenfilename(parent=root)
root.update()
root.destroy()

file_name = Path(file_path).name
fc = file_name.split("_")[0]
username = getpass.getuser()

data = pd.read_excel(file_path, header=None, sheet_name="Annex 4")
df = pd.DataFrame(data)

classification_xls = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FSI\classification.xlsx')
cls = pd.read_excel(classification_xls, sheet_name="MEMO", header=None)
df_cls = pd.DataFrame(cls)

codes_from_classification = df_cls.iloc[:, 4].to_list()
unit = [str(unit) for unit in df.iloc[:160, 5]] # changed


def nullcheck(value):
    if str(value) == "NaN" or str(value) == "nan":
        return None
    elif float(value) == 0:
        return None
    else:
        return str(value)


scale = 1000000
row = 12

cls_ids = []
dates = []
units =[]
values = []
created_time = []
created_by = []

# %Y-%m-%d %H:%M:%S"

# march 31
# june  30
# september 30
# december 31
# index_of_code = 2
# index_of_unit = 5
# index_of_dates = 9


dates_row = [row for row in df.iloc[9].values.tolist() if str(row) != "nan"]  #changed
date_full = [row for row in df.iloc[9].values.tolist()]  #changed

for i, row in df.iterrows():
        code = row.tolist()[2] # changed
        unit = row.tolist()[5]
        print(code)
        if str(code) != "nan" and code in codes_from_classification:
            print("ok")
            # now find the index of the code in the classification file
            index = codes_from_classification.index(str(code))
            # now match that index in the cls array
            for i, j in enumerate(row.tolist()):
                #TODO compile all into one file and index
                # i also changes
                if i > 5 and i <= df.iloc[9].values.tolist().index(str(dates_row[-1])): #changed
                    year = date_full[i][:4]
                    date = ""
                    if int(date_full[i][-1]) == 1:
                        date = year + "-03-31"
                    elif int(date_full[i][-1]) == 2:
                        date = year + "-06-30"
                    elif int(date_full[i][-1]) == 3:
                        date = year + "-09-30"
                    elif int(date_full[i][-1]) == 4:
                        date = year + "-12-31"



                    cls_ids.append(index)
                    dates.append(date)
                    units.append(unit)
                    if str(unit) == "Million" and nullcheck(j) != None:
                        values.append(float(nullcheck(j))*scale)
                    else:
                        values.append(nullcheck(j))

                    created_time.append(_datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    created_by.append(username)



dbdf = pd.DataFrame({
    'cls_id': cls_ids,
    'date': dates,
    'unit':units,
    'value': values,
    'created_by': created_by,
    'created_time': created_time
})

conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FSI;Trusted_Connection=yes;')

cursor = conn.cursor()
print("Deleting old data..")
tk_sql = "Delete from [FSI].[dbo].[F_FSI2A4];DBCC CHECKIDENT ('[FSI].[dbo].[F_FSI2A4]', RESEED, 0)"
cursor.execute(tk_sql)
cols = ",".join([str(i) for i in dbdf.columns.tolist()])


dbdf = dbdf.where(pd.notnull(dbdf), None)
for i, row in dbdf.iterrows():
    sql = "INSERT INTO [FSI].[dbo].F_FSI2A4 (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)"
    cursor.execute(sql, tuple(row))
    conn.commit()
print("Done!", i, 'rows updated!!')
