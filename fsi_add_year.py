import helpers
dbname = "FSI"
sheetname = "Annex 2"

db_name ={
    "Annex 2": "F_FSI2A2",
    "Annex 3": "F_FSI2A3",
    "Annex 4": "F_FSI2A4",
    "BS": "F_OFCB",
    "Table 1": "F_FSID"
}

class_names ={
    "Annex 2": "INCOME_EXPENSE",
    "Annex 3": "DEPOSIT",
    "Annex 4": "MEMO",
    "BS": "OFC",
    "Table 1": "FSID"
}

main_dataframe = helpers.getFile("Annex 2")
cls_dataframe = helpers.get_classification(r'\\admma2\data\SSD\SIS\FSI\classification.xlsx',class_names[sheetname])
codes_from_classification = cls_dataframe.iloc[:, 4].to_list()

# conn = helpers.makeConnection(dbname)

# print(helpers.checkDbf(dbname,db_name[sheetname],"2019-03-31"))

# helpers.truncate(dbname,db_name["Annex 2"])

dates_row = [row for row in main_dataframe.iloc[9].values.tolist() if str(row) != "nan"]  #changed
date_full = [row for row in main_dataframe.iloc[9].values.tolist()]  #changed

cls_ids =[]
dates = []
units =[]
values =[]
created_by =[]
created_time =[]

def getIndex(dates_full, year):
    index = []
    for i,j  in enumerate(date_full):
        if str(j) != "nan":
            if str(j)[:4] =="2015":
                index.append(i)
    return [index[0],index[-1]]

# helpers.remove_data(dbname,db_name["Annex 2"],"2015")

init_index = getIndex(date_full,"2015")

for i, row in main_dataframe.iterrows():
        code = row.tolist()[2] # changed
        unit = row.tolist()[5]
        print(code)
        if str(code) != "nan" and str(code) in codes_from_classification:
            print("ok")
            # now find the index of the code in the classification file
            index = codes_from_classification.index(str(code))
            # now match that index in the cls array
            for i, j in enumerate(row.tolist()):
                if i >= init_index[0] and i <= init_index[1]:
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
                    if str(unit) == "Million" and helpers.nullcheck(j) != None:
                        values.append(float(helpers.nullcheck(j))* 1000000)
                    else:
                        values.append(helpers.nullcheck(j))

                    created_time.append(helpers.datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    created_by.append(helpers.getpass.getuser())

dataframe = helpers.pd.DataFrame({
    'cls_id': cls_ids,
    'date': dates,
    'unit':units,
    'value': values,
    'created_by': created_by,
    'created_time': created_time
})


conn = helpers.makeConnection(dbname)
helpers.addToDb(dataframe,dbname,db_name["Annex 2"])
# helpers.truncate(dbname, db_name[sheetname],2015)
