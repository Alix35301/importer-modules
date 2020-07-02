from datetime import datetime
from pathlib import Path
from tkinter import filedialog
import getpass
import math
import numpy as np
import os
import pandas as pd
import pyodbc
import time
import tkinter as tk

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', 1)
file_path = filedialog.askopenfilename(parent=root)
root.update()
root.destroy()

file_name = Path(file_path).name

fc = file_name.split("_")[0]

data = pd.read_excel(file_path, header=None, sheet_name="FSID")
df = pd.DataFrame(data)
data = df.iloc[1:]

parent_id = []
description = []
unary_operator = []

for i, row in df.iterrows():
    if i > 0:
        if str(row.tolist()[1]) == "NaN" or str(row.tolist()[1]) == "nan":
            parent_id.append(None)
        else:
            parent_id.append(str(row.tolist()[1]))
        description.append(str(row.tolist()[2]))
        unary_operator.append(str(row.tolist()[3]))
data_to_db = pd.DataFrame({
    'parent_id': parent_id,
    'description': description,
    'unary_operator': unary_operator,
})

conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FSI;Trusted_Connection=yes;')

cursor = conn.cursor()

# to truncate the table uncomment this line
tk_sql = "Delete from [FSI].[dbo].[D_FSID];DBCC CHECKIDENT ('[FSI].[dbo].[D_FSID]', RESEED, 0)"
cursor.execute(tk_sql)
cols = ",".join([str(i) for i in data_to_db.columns.tolist()])

for i, row in data_to_db.iterrows():
    sql = "INSERT INTO [FSI].[dbo].D_FSID (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)"
    cursor.execute(sql, tuple(row))

    # the connection is not autocommitted by default, so we must commit to save our changes
    conn.commit()
print("Classfication table is updated!")