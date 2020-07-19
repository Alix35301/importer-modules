from datetime import datetime
import getpass
from pathlib import Path
from pandas.io import sql
import helpers
import MySQLdb


def connectToDB():
    # conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=LOANS;Trusted_Connection=yes;')
    conn = MySQLdb.connect(host="localhost",
                           user="root",
                           passwd="1234",
                           db="tourism")
    return conn


base_excel_file = helpers.pd.ExcelFile("QNA-2019-Q4-Tables.xlsx")
gdp_df = helpers.pd.read_excel(base_excel_file, sheet_name="Table 1", header=None)

classification_xls = helpers.pd.ExcelFile("gdp table data.xlsx")
GDP_SECTORS = helpers.pd.read_excel(classification_xls, sheet_name="D_GDP_SECTOR", header=None,).iloc[:,2].to_list()
GDP_SECTORS_ID = helpers.pd.read_excel(classification_xls, sheet_name="D_GDP_SECTOR", header=None,).iloc[:,0].to_list()

YEARS = gdp_df.iloc[4].to_list()
QT = gdp_df.iloc[5].to_list()


date =[]
sector_id =[]
amount =[]
date_entered =[]
user_name = []
freq = []


for i in YEARS:
     if str(i) != "nan" and YEARS.index(i) > 1:
         index = YEARS.index(i)
         for j in range(3):
             index =index+1
             YEARS[index]=i
         # index=0

date_L = []
sector_L = []
old_V = []
new_V =[]

def duplicated(sector_id, time, amount):
    conn = connectToDB()
    sql = f"SELECT * FROM tourism.F_C_GDP_Q where SECTOR_ID = {sector_id} and DATE_TIME = '{time}' "
    # print(sql)
    cursor = conn.cursor()
    cursor.execute(sql)
    result = cursor.fetchall()
    if len(result) > 0:
        if amount == float(result[0][3]):
            # print(f"{amount} == {result[0][3]}")
            # print("Skipping")
            return True
        else:
            sql_update = f"UPDATE tourism.F_C_GDP_Q SET GDP = {amount} where DATE_TIME = '{time}' and SECTOR_ID = {sector_id} "
            cursor.execute(sql_update)
            print(sql_update)
            conn.commit()
            date_L.append(time)
            sector_L.append(sector_id)
            old_V.append(result[0][3])
            new_V.append(amount)
            return True
    else:
        return False
        # skipping

def getSector(secid):
    sector = None
    for i in GDP_SECTORS_ID:
        if i == secid:
            sector = GDP_SECTORS[GDP_SECTORS_ID.index(i)]
            break
    return sector


for i,j in gdp_df.iterrows():
    if i > 5:
        rowdata = j.to_list()
        # print(rowdata[1])
        if str(rowdata[1]).strip() in GDP_SECTORS:
            id =GDP_SECTORS_ID[GDP_SECTORS.index(str(rowdata[1]).strip())]
            # print(id)
            for l, m in enumerate(rowdata):
                if l > 1 and str(m) != "nan":
                    year = YEARS[l]
                    d = None
                    if QT[l] =="Q1":
                        d = f"{year}-03-31"

                    elif QT[l] == "Q2":
                        d = f"{year}-06-30"

                    elif QT[l] == "Q3":
                        d = f"{year}-09-30"
                    elif QT[l] == "Q4":
                        d = f"{year}-12-31"
                    if not duplicated(id,d,m):
                        sector_id.append(id)
                        amount.append(m)
                        date.append(d)
                        date_entered.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        user_name.append(getpass.getuser())
                        freq.append("quarterly")

dbdf = helpers.pd.DataFrame({
"DATE_TIME":date,
"SECTOR_ID":sector_id,
"GDP":amount,
"DATE_ENTERED":date_entered,
"USER_NAME":user_name,
"frequency":freq

})

helpers.addToDb(dbdf,"tourism","F_C_GDP_Q")


if len(date_L)>0:
    df_Log = helpers.pd.DataFrame({
        "DATE":date_L,
        "SECTOR":sector_L,
        "OLD_VAUE":old_V,
        "NEW_VALUE":new_V
    })

    file_name = f"GDP-updatelog-{datetime.now().strftime('%Y-%m-%d %H%M%S')}.csv"

    df_Log.to_csv(file_name, index=False, header=True)
