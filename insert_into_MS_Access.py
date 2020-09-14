# September 13, 2020

import pyodbc as pyo

print("opening the data connection...")
cnn_string = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\Home\Desktop\pyintegrate.accdb" # my local dbase file
    )
cnn = pyo.connect(cnn_string)
cursor = cnn.cursor()
sql = (
    r"Insert into PROJECT_DATA (InfoType, Comment) Values" # tbl name and column headers to insert into
    r"('Meeting', 'Review Meeting scheduled for Thursday')" # data to insert
    )
cursor.execute(sql)
cursor.commit()
sql = "select * from PROJECT_DATA"
cursor.execute(sql)
for row in cursor.fetchall():
    print(row)
print("insert successful...")
cursor.close()
cnn.close()
print("data connection closed...")
