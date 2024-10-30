import pyodbc

con_string = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=G:\Shared drives\Root\11. DATABASE\DBRGT-02.accdb;"
)

try:
    conn = pyodbc.connect(con_string)
    print("Connection successful")
except pyodbc.Error as ex:
    sqlstate = ex.args[0]
    print("Connection failed: ", sqlstate)
