import pyodbc
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
import getpass
import pandas as pd
import pyodbc
import tkinter as tk

def add_yearly_data(y):

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(parent=root)
    root.update()
    root.destroy()

    file_name = Path(file_path).name
    fc = file_name.split("_")[0]
    username = getpass.getuser()

    data = pd.read_excel(file_path, header=None, sheet_name="FASurvey")
    df = pd.DataFrame(data)

    years = df.iloc[9, :].values.tolist()
    col = years.index(y) # pass year to be replaced

    # reading from old format
    old_format = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FAS\formats\old format.xlsx')
    old_format_pd = pd.read_excel(old_format, sheet_name="FASurvey", header=None)
    # data frame of old format
    df_old = pd.DataFrame(old_format_pd)
    colDataOld = df_old.iloc[:218, col]
    old_codes = [str(code).strip() for code in df_old.iloc[:218, 2]]
    old_unit = [str(unit) for unit in df_old.iloc[:218, 3]]

    classification_xls = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FAS\classification_updated.back.xlsx')
    cls = pd.read_excel(classification_xls, sheet_name="cls", header=None)
    clsIdArray = cls.loc[1:, 0].values.tolist()
    maping = cls.loc[1:, 4].values.tolist()

    checkValue = lambda v: None if v == 'nan' or v == '' else int(float(v))

    # based on the column index
    colData = df.iloc[:160, col]
    code = [str(code).strip() for code in df.iloc[:160, 2]]
    unit = [str(unit) for unit in df.iloc[:160, 3]]
    datatodb=[]
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

    for index, id in enumerate(clsIdArray):
        l = []  # list to place comma-seperated codes found in classification
        if str(maping[index]) != 'nan':
            l = maping[index].split(",")
        else:
            continue
        for i in l:
            if i in code:
                index = code.index(i)
                data = colData[code.index(i)]
                u = unit[code.index(i)]
                if index > 105 and str(colData[code.index(i)]) != 'nan':
                    if unit[code.index(i)] == 'Million':
                        l[l.index(i)] = str(float(colData[code.index(i)]) * 1000000)
                    else:
                        l[l.index(i)] = str(colData[code.index(i)])
                else:
                    l[l.index(i)] = str(colData[code.index(i)])
            else:
                # time to read old format
                if i in old_codes:

                    index = old_codes.index(i)
                    data = colDataOld[old_codes.index(i)]
                    if index > 105 and str(colDataOld[old_codes.index(i)]) != 'nan':
                        if old_unit[old_codes.index(i)] == 'Million':
                            l[l.index(i)] = str(float(colDataOld[old_codes.index(i)]) * 1000000)
                        else:
                            l[l.index(i)] = str(colDataOld[old_codes.index(i)])
                        pass
                    else:
                        l[l.index(i)] = str(colDataOld[old_codes.index(i)])

        if check(l) != None:
            l.insert(0, str(id))
            l.insert(1, str(y))
            datatodb.append(l)


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
    remove_data(y)
    add_data_to_db(dbdf)
    print("{} Successfully added to the database".format(y))

def remove_data(year):
    conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
    cursor = conn.cursor()
    sql = ("""DELETE FROM [FAS].[dbo].DATA WHERE year = '%s'""" % year)
    cursor.execute(sql)

def add_data_to_db(dbdf):
    dbdf = dbdf.where(pd.notnull(dbdf), None)
    for i, row in dbdf.iterrows():
        cols = ",".join([str(i) for i in dbdf.columns.tolist()])
        conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
        cursor = conn.cursor()
        sql = "INSERT INTO [FAS].[dbo].DATA (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)"
        cursor.execute(sql, tuple(row))
        conn.commit()

def check(l):
    count = 1
    for i in l:
        if len(i) < 1 or i == 'nan':
            count += 1
    if count > len(l):
        return
    else:
        return l


def checkDb(year):
    conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
    cursor = conn.cursor()
    sql = ("""SELECT * FROM [FAS].[dbo].DATA WHERE year = '%s'""" % year)
    cursor.execute(sql)
    result = cursor.fetchall()
    if len(result)<1:
        print("Data not found.. adding now!")
        add_yearly_data(year)
    else:
        response = input("Data already exists! Would you like to overwrite the data ? (Y/N)")
        if response.upper() == "Y":
            add_yearly_data(year)
            print("Overwriting data..")
        elif response.upper() =="N":
            print("Ok")

print("Options")
print("1. Bulk import")
print("2. Add year")
x=input()

if int(x)==1:
    #run main fas importer
    pass
elif int(x)==2:
    year = input("Enter year to be imported! ")
    #check if the year exists in the database
    checkDb(year)


