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


base_excel_file = helpers.pd.ExcelFile("Apr 2019.xlsx")
arrivals_df = helpers.pd.read_excel(base_excel_file, sheet_name="T1", header=None)

classification_xls = helpers.pd.ExcelFile("t.xlsx")
df_d_country = helpers.pd.read_excel(classification_xls, sheet_name="d country classification", header=None)
df_d_time = helpers.pd.read_excel(classification_xls, sheet_name="d tourism time", header=None)

country_codes = df_d_country.iloc[:, 4].tolist()
country_Id = df_d_country.iloc[:, 0].tolist()
country_codes = [str(i).strip() for i in country_codes]

d_time_years = df_d_time.iloc[:, 4].tolist()
d_time_calender_months = df_d_time.iloc[:, 2].tolist()
d_time_ids = df_d_time.iloc[:, 0].tolist()

time_id_L =[]
countryCode_L =[]
old_L =[]
new_L = []

def getCountry(id):
    country = None
    for i in country_Id:
        if id == i:
            country = country_codes[country_Id.index(id)]
            break
    return country

def getTimeID(month, year):
    timeid = 0
    for i, id in enumerate(d_time_ids):
        if d_time_calender_months[i] == month and d_time_years[i] == year:
            timeid = id
    return timeid


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


def duplicated(timeid, countryId, amount):
    conn = connectToDB()
    sql = f"SELECT * FROM tourism.F_TOURISM_ARRIVALS where TIME_ID = {timeid} and COUNTRY_ID = {countryId} "
    cursor = conn.cursor()
    cursor.execute(sql)
    result = cursor.fetchall()
    if len(result) > 0:
        if amount == result[0][3]:
            # print(f"{amount} == {result[0][3]}")
            # print("Skipping")
            return True
        else:
            sql_update = f"UPDATE tourism.F_TOURISM_ARRIVALS SET AMOUNT = {amount} where TIME_ID = {timeid} and COUNTRY_ID = {countryId} "
            cursor.execute(sql_update)
            conn.commit()
            time_id_L.append(timeid)
            countryCode_L.append(getCountry(countryId))
            old_L.append(result[0][3])
            new_L.append(amount)
            return True
    else:
        return False
        # skipping


meta = getMetaData(base_excel_file)

currentYear = getTimeID(datetime.strptime(meta.get('month'), '%b').strftime('%B'), int(meta.get('year')))
previousYear = getTimeID(datetime.strptime(meta.get('month'), '%b').strftime('%B'), int(meta.get('year')) - 1)

# this is done to make it easier to grab the time id whle looping

daterow = arrivals_df.iloc[3].tolist()
daterow[3] = currentYear
daterow[2] = previousYear

country_code_list = []
time_id_list = []
amount_list = []
arrival_type_list = []
changed_list = []

country_alt_names = {
    "Vietnam": ["Viet Nam"],
    "STATELESS/NOT STATED": ["Other / Not Stated"],
    "Russia":["Bromia"]
}
country_codes_from_main = arrivals_df.iloc[:, 1].tolist()
country_regions_from_main = arrivals_df.iloc[:, 0].tolist()

country_not_present = []

for i, j in enumerate(country_codes_from_main):
    if str(j).strip() != "nan" and str(j).strip() not in country_codes:
        for key, val in country_alt_names.items():
            if str(j).strip() in val:
                country_codes_from_main[i] = key
                break
        else:
            country_not_present.append(j)
if len(country_not_present) > 0:
    print("WARNING!  Could Not find the following countries in the classification.. "
          "\nConsider adding or set alternative names")
    [print(i) for i in country_not_present]
    exit(0)



for i, j in arrivals_df.iterrows():
    rowdata = j.tolist()
    # print(country_codes_from_main[i])
    if str(country_codes_from_main[i]).strip() in country_codes:
        for l, m in enumerate(rowdata):
            if l == 2:
                if not duplicated(int(daterow[2]), country_Id[country_codes.index(str(country_codes_from_main[i]).strip())], rowdata[l]):
                    time_id_list.append(int(daterow[2]))
                    country_code_list.append(country_Id[country_codes.index(str(country_codes_from_main[i]).strip())])
                    amount_list.append(rowdata[l])
                    arrival_type_list.append(2)
                    changed_list.append(None)
            elif l == 3:

                if not duplicated(int(daterow[3]), country_Id[country_codes.index(str(country_codes_from_main[i]).strip())], rowdata[l]):
                    time_id_list.append(int(daterow[3]))
                    country_code_list.append(country_Id[country_codes.index(str(country_codes_from_main[i]).strip())])
                    amount_list.append(rowdata[l])
                    arrival_type_list.append(2)
                    changed_list.append(None)

for i, j in arrivals_df.iterrows():
    rowdata = j.tolist()
    if str(country_regions_from_main[i]).strip() in country_codes:
        for l, m in enumerate(rowdata):
            if l == 2:
                if not duplicated(int(daterow[2]), country_Id[country_codes.index(str(country_regions_from_main[i]).strip())], rowdata[l]):
                    time_id_list.append(int(daterow[2]))
                    country_code_list.append(country_Id[country_codes.index(str(country_regions_from_main[i]).strip())])
                    amount_list.append(rowdata[l])
                    arrival_type_list.append(2)
                    changed_list.append(None)
            elif l == 3:

                if not duplicated(int(daterow[3]), country_Id[country_codes.index(str(country_regions_from_main[i]).strip())], rowdata[l]):
                    time_id_list.append(int(daterow[3]))
                    country_code_list.append(country_Id[country_codes.index(str(country_regions_from_main[i]).strip())])
                    amount_list.append(rowdata[l])
                    arrival_type_list.append(2)
                    changed_list.append(None)

dbdf = helpers.pd.DataFrame({
    "TIME_ID": time_id_list,
    "COUNTRY_ID": country_code_list,
    "AMOUNT": amount_list,
    "ARRIVAL_TYPE_ID": arrival_type_list,
    "CHANGED": changed_list

})
# add time to d time table

helpers.addToDb(dbdf, "tourism", "F_TOURISM_ARRIVALS")


if len(time_id_L)>0:
    df_Log = helpers.pd.DataFrame({
        "TIME_ID":time_id_L,
        "COUNTRY_CODE":countryCode_L,
        "OLD_VAUE":old_L,
        "NEW_VALUE":new_L
    })

    file_name = f"Arrivals_{meta.get('filename')[:8]}-updatelog.csv"

    df_Log.to_csv(file_name, index=False, header=True)
