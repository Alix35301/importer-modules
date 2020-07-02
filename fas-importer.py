import helpers
import getpass
import pandas as pd
import pyodbc

username = getpass.getuser()

file_path = helpers.getFile()  # reads the main data file
data = pd.read_excel(file_path, header=None, sheet_name="FASurvey")
df = pd.DataFrame(data)

# Function call to extract the dates in the spread sheet
getdates = lambda: [year for year in df.iloc[9, :] if str(year) != 'nan']
checkValue = lambda v: None if v == 'nan' or v == '' else int(float(v))

# def updateColumnData(c):
#     colData.clear()
#     [colData.append(data) for data in df.iloc[:160, c]]
#     colDataOld.clear()
#     [colDataOld.append(data) for data in old_format_pd.iloc[:218, c]]


# def updatecoldata(c):
#     colData.clear()
#     return [colData.append(data) for data in df.iloc[:160, c]]
#
#
#
# def updateOldcoldata(c):
#     # for y in getdates():
#     # store these in array
#     colDataOld.clear()
#     return [colDataOld.append(data) for data in old_format_pd.iloc[:218, c]]
#     # for i in df.iloc[:160, c]:
#     #     colData.append(i)
#     # return colData


# Loading cls file to get parent ids
classification_xls = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FAS\classification_updated.back.xlsx')
cls = pd.read_excel(classification_xls, sheet_name="cls", header=None)
clsIdArray = cls.loc[1:, 0].values.tolist()
maping = cls.loc[1:, 4].values.tolist()

# reading from old format
old_format = pd.ExcelFile(r'\\admma2\data\SSD\SIS\FAS\formats\old format.xlsx')
old_format_pd = pd.read_excel(old_format, sheet_name="FASurvey", header=None)
# data frame of old format
df_old = pd.DataFrame(old_format_pd)
old_codes = [str(code).strip() for code in df_old.iloc[:218, 2]]

code = []
unit = []
colData = []
colDataOld = []
datatodb = []

# cls_id = []
# year = []
# instiution = []
# branches = []
# non_branches = []
# atm = []
# dpst_ins_holder = []
# dpst_acct_ins_plc = []
# borrower = []
# loan_acct = []
# cards = []
# outsd_ins_tech_res = []
# outsd_loan = []
# mobile_inet = []
# mobile_money = []
# created_by = []
# created_date = []
unit = [str(unit) for unit in df.iloc[:160, 3]]
old_unit = [str(unit) for unit in df.iloc[:218, 3]]
code = [str(code).strip() for code in df.iloc[:160, 2]]
c = 5
for y in getdates():
    c += 1
    helpers.updateColumnData(c,colData,colDataOld,df,old_format_pd)
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

                if index > 105 and colData[code.index(i)] != 'nan':
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
                    if index > 105 and colDataOld[old_codes.index(i) != 'nan']:
                        if old_unit[old_codes.index(i)] == 'Million':
                            l[l.index(i)] = str(float(colDataOld[old_codes.index(i)]) * 1000000)
                        else:
                            l[l.index(i)] = str(colDataOld[old_codes.index(i)])
                    else:
                        l[l.index(i)] = str(colDataOld[old_codes.index(i)])

        if helpers.check(l) != None:
            l.insert(0, str(id))
            l.insert(1, str(y))
            datatodb.append(l)

# adding collected data to lists to make dataframe for sql input

dbdf = helpers.add_data_to_dataframe(datatodb)
# for data in datatodb:
#     for index, item in enumerate(data):
#         if index == 0:
#             cls_id.append(checkValue(item))
#         elif index == 1:
#             year.append(checkValue(item))
#         elif index == 2:
#             instiution.append(checkValue(item))
#         elif index == 3:
#             branches.append(checkValue(item))
#         elif index == 4:
#             non_branches.append(checkValue(item))
#         elif index == 5:
#             atm.append(checkValue(item))
#         elif index == 6:
#             dpst_ins_holder.append(checkValue(item))
#         elif index == 7:
#             dpst_acct_ins_plc.append(checkValue(item))
#         elif index == 8:
#             borrower.append(checkValue(item))
#         elif index == 9:
#             loan_acct.append(checkValue(item))
#         elif index == 10:
#             cards.append(checkValue(item))
#         elif index == 11:
#             outsd_ins_tech_res.append(checkValue(item))
#         elif index == 12:
#             outsd_loan.append(checkValue(item))
#         elif index == 13:
#             mobile_inet.append(checkValue(item))
#         elif index == 14:
#             mobile_money.append(checkValue(item))
#         elif index == 15:
#             created_by.append(username)
#         elif index == 16:
#             created_date.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
# dbdf = pd.DataFrame({
#     'cls_id': cls_id,
#     'year': year,
#     'institution': instiution,
#     'branches': branches,
#     'non_branches': non_branches,
#     'atm': atm,
#     'dpst_ins_holder': dpst_ins_holder,
#     'dpst_acct_ins_plc': dpst_acct_ins_plc,
#     'borrower': borrower,
#     'loan_acct': loan_acct,
#     'cards': cards,
#     'outsd_ins_tech_res': outsd_ins_tech_res,
#     'outsd_loan': outsd_loan,
#     'mobile_inet': mobile_inet,
#     'mobile_money': mobile_money,
#     'created_by': created_by,
#     'created_time': created_date
# })

# conn = pyodbc.connect('Driver={SQL Server};Server=noo\\noo;Database=FAS;Trusted_Connection=yes;')
#
# cursor = conn.cursor()
# print("Deleting old data..")
# tk_sql = ("""Delete from [FAS].[dbo].[DATA];DBCC CHECKIDENT ('[FAS].[dbo].[DATA]', RESEED, 0)""")
# cursor.execute(tk_sql)
# cols = ",".join([str(i) for i in dbdf.columns.tolist()])
#
# dbdf = dbdf.where(pd.notnull(dbdf), None)
# for i, row in dbdf.iterrows():
#     sql = "INSERT INTO [FAS].[dbo].DATA (" + cols + ") VALUES (" + "?," * (len(row) - 1) + "?)"
#     cursor.execute(sql, tuple(row))
#     conn.commit()
# print("Done!", i, 'rows updated!!')
