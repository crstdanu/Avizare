import os
import functii as x
import pyodbc
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta



id_executie = 1

con_string = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:/Shared drives/Root/11. DATABASE/RoGoTehnic-DataBase-2024.accdb;"
conn = pyodbc.connect(con_string)
cursor = conn.cursor()

def fetch_single_value(cursor, query, params):
    return cursor.execute(query, params).fetchval()

ce_date_sunt =  fetch_single_value(cursor, 'SELECT ValoareAC from tblIncepereExecutie WHERE ID_Lucrare = ?', (id_executie,))
print(float(round(ce_date_sunt, 2))*0.01)

# data_asta = x.get_date(ce_date_sunt)

# new_date = ce_date_sunt + relativedelta(months=6)
# formatted_date = new_date.strftime('%d-%m-%Y')

# print(formatted_date)
