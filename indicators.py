from datetime import datetime
from pathlib import Path
from tkinter import filedialog

import MySQLdb
from xlrd import XLRDError
import ctypes
import getpass
import hashlib
import math
# import MySQLdb
import numpy as np
import os
import pandas as pd
import pyodbc
import time
import tkinter as tk
from time import strptime
import helpers

classification_xls = pd.ExcelFile("t.xlsx")
df_d_country = pd.read_excel(classification_xls, sheet_name="d country classification", header=None)
df_d_time = pd.read_excel(classification_xls, sheet_name="d tourism time", header=None)

country_codes = df_d_country.iloc[:, 4].tolist()
country_Id = df_d_country.iloc[:, 0].tolist()
country_codes = [str(i).strip() for i in country_codes]

d_time_years = df_d_time.iloc[:, 4].tolist()
d_time_calender_months = df_d_time.iloc[:, 2].tolist()
d_time_ids = df_d_time.iloc[:, 0].tolist()

# for logger
timeid_L=[]
rtype_L=[]
name_L=[]
oldV_L =[]
newV_L=[]

def connectToDB():
    # conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=LOANS;Trusted_Connection=yes;')
    conn = MySQLdb.connect(host="localhost",
                           user="root",
                           passwd="1234",
                           db="tourism")
    return conn


def checkDuplicates(timeid,rTypeId, reg_bed, op_bed, bed_night, est_op, est_reg):
    conn = connectToDB()
    sql = f"SELECT * FROM tourism.F_TOURISM_INDICATORS where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId}"
    cursor = conn.cursor()
    cursor.execute(sql)
    result = cursor.fetchall()
    if len(result) > 0:
        if round(float(reg_bed),4) == float(result[0][4]) or round(float(op_bed),4) == float(result[0][5]) or round(float(bed_night),4) == float(result[0][6]) or round(float(est_reg),4) == float(result[0][10]) or round(float(est_op),4) == float(result[0][9]):
            # print(f"{amount} == {result[0][3]}")
            print("Skipping")
            return True
        elif round(float(reg_bed),4) != float(result[0][4]):
            sql_update = f"UPDATE tourism.F_TOURISM_INDICATORS SET REGISTERED_BED_CAPACITY = {round(float(reg_bed),4)} where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId} "
            cursor.execute(sql_update)
            conn.commit()
            print("UPDATING RECORDS")
            timeid_L.append(timeid)
            rtype_L.append(rType)
            name_L.append("REGISTERED_BED_CAPACITY")
            oldV_L.append(float(result[0][4]))
            newV_L.append(round(float(reg_bed),4))
            return True
        elif round(float(op_bed),4) != float(result[0][5]):
            sql_update = f"UPDATE tourism.F_TOURISM_INDICATORS SET OPERATIONAL_BED_CAPACITY = {round(float(op_bed),4)} where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId} "
            cursor.execute(sql_update)
            conn.commit()
            timeid_L.append(timeid)
            rtype_L.append(rType)
            name_L.append("OPERATIONAL_BED_CAPACITY")
            oldV_L.append(float(result[0][4]))
            newV_L.append(round(float(reg_bed),4))
            return True
        elif round(float(bed_night),4) != float(result[0][6]):
            sql_update = f"UPDATE tourism.F_TOURISM_INDICATORS SET BED_NIGHTS = {round(float(bed_night),4)} where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId} "
            cursor.execute(sql_update)
            conn.commit()
            timeid_L.append(timeid)
            rtype_L.append(rType)
            name_L.append("BED_NIGHTS")
            oldV_L.append(float(result[0][4]))
            newV_L.append(round(float(reg_bed),4))
            return True
        elif round(float(est_reg),4) != float(result[0][10]):
            sql_update = f"UPDATE tourism.F_TOURISM_INDICATORS SET REGISTERED_NOS = {round(float(est_reg),4)} where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId} "
            cursor.execute(sql_update)
            conn.commit()
            timeid_L.append(timeid)
            rtype_L.append(rType)
            name_L.append("REGISTERED_NOS")
            oldV_L.append(float(result[0][4]))
            newV_L.append(round(float(reg_bed),4))
            return True

        elif round(float(est_op),4) != float(result[0][9]):
            sql_update = f"UPDATE tourism.F_TOURISM_INDICATORS SET OPERATIONAL_NOS = {round(float(est_op),4)} where TIME_ID = {timeid} and RESORT_TYPE_ID = {rTypeId} "
            cursor.execute(sql_update)
            conn.commit()
            timeid_L.append(timeid)
            rtype_L.append(rType)
            name_L.append("OPERATIONAL_NOS")
            oldV_L.append(float(result[0][4]))
            newV_L.append(round(float(reg_bed),4))
            return True

    else:
        return False


def getTimeID(month, year):
    timeid = 0
    for i, id in enumerate(d_time_ids):
        if d_time_calender_months[i] == month and d_time_years[i] == year:
            timeid = id
    return timeid


# connect to db


# open file dialog to choose the file
def selectFile():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)

    file_path = filedialog.askopenfilename(parent=root)

    root.update()
    root.destroy()

    return file_path


# get meta data of file and user
def getMetaData(file_path):
    file_name = Path(file_path).name
    success = True
    month = year = date = message = ''
    # return fc
    try:
        month = file_name.split(" ")[0]
        year = file_name.split(" ")[1].split('.')[0]
        date = datetime.strptime(f'{year}-{month}-01', "%Y-%b-%d").strftime("%Y-%m-%d")
    except (IndexError, ValueError):
        success = False
        message = "Invalid file selected. File name should be in the format (Jan 2019.xlsx)"

    username = getpass.getuser()
    output = {
        'success': success,
        'message': message,
        'month': month,
        'year': year,
        'date': date,
        'username': username,
        'filename':file_name
    }
    return output


# read and process file
def processFile(file_path, meta):
    try:
        data = pd.read_excel(file_path, header=None, sheet_name="T3-T5")
        df = pd.DataFrame(data)

        month = meta.get('month')

        if month == 'Jan':
            dset_one_start = 3
            dset_one_end = 3
            dset_two_start = 10
            dset_two_end = 10
            dset_three_start = 17
            dset_three_end = 17
        elif month == 'Feb':
            dset_one_start = 3
            dset_one_end = 4
            dset_two_start = 11
            dset_two_end = 12
            dset_three_start = 19
            dset_three_end = 20
        elif month == 'Mar':
            dset_one_start = 3
            dset_one_end = 5
            dset_two_start = 12
            dset_two_end = 14
            dset_three_start = 21
            dset_three_end = 23
        elif month == 'Apr':
            dset_one_start = 3
            dset_one_end = 6
            dset_two_start = 13
            dset_two_end = 16
            dset_three_start = 23
            dset_three_end = 26
        elif month == 'May':
            dset_one_start = 3
            dset_one_end = 7
            dset_two_start = 14
            dset_two_end = 18
            dset_three_start = 25
            dset_three_end = 29
        elif month == 'Jun':
            dset_one_start = 3
            dset_one_end = 8
            dset_two_start = 15
            dset_two_end = 20
            dset_three_start = 27
            dset_three_end = 32
        elif month == 'Jul':
            dset_one_start = 3
            dset_one_end = 9
            dset_two_start = 21
            dset_two_end = 27
            dset_three_start = 39
            dset_three_end = 45
        elif month == 'Aug':
            dset_one_start = 3
            dset_one_end = 10
            dset_two_start = 17
            dset_two_end = 24
            dset_three_start = 31
            dset_three_end = 38
        elif month == 'Sep':
            dset_one_start = 3
            dset_one_end = 11
            dset_two_start = 18
            dset_two_end = 26
            dset_three_start = 33
            dset_three_end = 41
        elif month == 'Oct':
            dset_one_start = 3
            dset_one_end = 12
            dset_two_start = 19
            dset_two_end = 28
            dset_three_start = 35
            dset_three_end = 44
        elif month == 'Nov':
            dset_one_start = 3
            dset_one_end = 13
            dset_two_start = 20
            dset_two_end = 30
            dset_three_start = 37
            dset_three_end = 47
        elif month == 'Dec':
            dset_one_start = 3
            dset_one_end = 14
            dset_two_start = 21
            dset_two_end = 32
            dset_three_start = 39
            dset_three_end = 50

        est_reg_start = 1
        est_reg_end = 5
        est_op_start = 6
        est_op_end = 10
        bed_reg_start = 1
        bed_reg_end = 5
        bed_op_start = 6
        bed_op_end = 10
        bed_night_start = 1
        bed_night_end = 5

        estRegDF = df.loc[dset_one_start:dset_one_end, est_reg_start:est_reg_end]
        estRegDF = estRegDF.reset_index(drop=True)
        estRegDF = estRegDF.T.reset_index(drop=True).T
        estOpDF = df.loc[dset_one_start:dset_one_end, est_op_start:est_op_end]
        estOpDF = estOpDF.reset_index(drop=True)
        estOpDF = estOpDF.T.reset_index(drop=True).T
        bedRegDF = df.loc[dset_two_start:dset_two_end, bed_reg_start:bed_reg_end]
        bedRegDF = bedRegDF.reset_index(drop=True)
        bedRegDF = bedRegDF.T.reset_index(drop=True).T
        bedOpDF = df.loc[dset_two_start:dset_two_end, bed_op_start:bed_op_end]
        bedOpDF = bedOpDF.reset_index(drop=True)
        bedOpDF = bedOpDF.T.reset_index(drop=True).T
        bedNightDF = df.loc[dset_three_start:dset_three_end, bed_night_start:bed_night_end]
        bedNightDF = bedNightDF.reset_index(drop=True)
        bedNightDF = bedNightDF.T.reset_index(drop=True).T

        output = {
            'success': True,
            'message': 'Data processing successful!',
            'est_reg': estRegDF,
            'est_op': estOpDF,
            'bed_reg': bedRegDF,
            'bed_op': bedOpDF,
            'bed_night': bedNightDF,
        }
        return output
    except (ValueError, XLRDError) as e:
        message = 'T1 sheet not found in selected file'
        output = {
            'success': False,
            'message': message
        }
        return output


file = selectFile()
op = processFile(file, getMetaData(file))
est_reg = op.get("est_reg")
est_op = op.get("est_op")
bed_reg = op.get("bed_reg")
bed_op = op.get("bed_op")
bed_night = op.get("bed_night")

r, c = est_reg.shape

resort_typeId = []
timeId = []
arrival_type =[]
est_reg_list = []
est_op_list = []
bed_reg_list = []
bed_op_list = []
bed_night_list = []
avg_stay = []
ocu_rate = []
bed_night_cap = []
created_date = []
created_by = []

month_number = "1"
meta = getMetaData(file)
rType = None
for y in range(r):
    datetime_object = datetime.strptime(month_number, "%m")
    month_name = datetime_object.strftime("%B")

    month_number = str(int(month_number) + 1)
    for x in range(c):
        if x == 0:
            rType =1
        elif x == 1:
            rType = 2
        elif x == 2:
            rType = 3
        elif x == 3:
            rType = 4
        elif x == 4:
            rType = 5
        tid =getTimeID(month_name, int(meta.get("year")))
        if not checkDuplicates(tid,rType,bed_reg.iloc[y, x],bed_op.iloc[y, x],bed_night.iloc[y, x],est_op.iloc[y, x],est_reg.iloc[y, x]):
            arrival_type.append(1)
            resort_typeId.append(rType)
            timeId.append(tid)
            est_reg_list.append(round(float(est_reg.iloc[y, x]),4))
            est_op_list.append(round(float(est_op.iloc[y, x]),4))
            bed_reg_list.append(round(float(bed_reg.iloc[y, x]),4))
            bed_op_list.append(round(float(bed_op.iloc[y, x]),4))
            bed_night_list.append(round(float(bed_night.iloc[y, x]),4))
            avg_stay.append(1)
            ocu_rate.append(1)
            bed_night_cap.append(1)
            created_by.append(getpass.getuser())
            created_date.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

dbdf = pd.DataFrame({
    "RESORT_TYPE_ID":resort_typeId,
    "ARRIVAL_TYPE_ID": arrival_type,
    "TIME_ID": timeId,
    "REGISTERED_BED_CAPACITY": bed_reg_list,
    "OPERATIONAL_BED_CAPACITY": bed_op_list,
    "BED_NIGHTS": bed_night_list,
    "OCCUPANCY_RATE": ocu_rate,
    "AVG_STAY_DURATION": avg_stay,
    "REGISTERED_NOS": est_reg_list,
    "OPERATIONAL_NOS": est_op_list,
    "BED_NIGH_CAPACITY": bed_night_cap,
    "ENTERED_DATE": created_date,
    "STAFF_USER": created_by
})

helpers.addToDb(dbdf,"tourism","F_TOURISM_INDICATORS")

if len(timeid_L)>0:
    df_Log = pd.DataFrame({
        "TIME_ID":timeid_L,
        "RESORT_TYPE_ID":rtype_L,
        "NAME":name_L,
        "OLD_VALUE":oldV_L,
        "NEW_VALUE":newV_L
    })

    file_name = f"{meta.get('filename')[:8]}-updatelog.csv"

    df_Log.to_csv(file_name, index=False, header=True)

print("done")
