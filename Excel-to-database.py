import xlrd, sqlite3
from tkinter import filedialog
from tkinter import *
root = Tk()
bol = True
while bol:
    file_location = filedialog.askopenfilename(title = "Select file", filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    if file_location != "":
        bol = False
root.destroy()
workbook = xlrd.open_workbook(file_location)
name = file_location[file_location.rfind("/")+1:-5].replace(" ", "_")
connection = sqlite3.connect(name + ".db")
crsr = connection.cursor()
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(row, col) for col in range(sheet.ncols)] for row in range(sheet.nrows)]
reverse_data = [[sheet.cell_value(row, col) for row in range(sheet.nrows)] for col in range(sheet.ncols)]
types = []
for i in range(len(reverse_data)):
    if all(isinstance(x, (int, float)) for x in reverse_data[i][1:]):
        types.append('INTEGER')
    else:
        types.append('TEXT')
try:
    sql_command = """SELECT * FROM table_""" + name
    crsr.execute(sql_command)
except:
    sql_command = """CREATE TABLE table_""" + name + """ ("""
    for i in range(len(types)):
        sql_command += data[0][i]
        sql_command += ' ' + types[i] + ', '
    sql_command = sql_command[:-2]
    sql_command += ')'
    crsr.execute(sql_command)
for j in range(len(data[1:])):
    sql_command = """INSERT INTO table_""" + name + """  VALUES ("""
    sql_command += str(data[j+1])[1:-1]
    sql_command += ')'
    crsr.execute(sql_command)
sql_command = """SELECT * FROM table_""" + name
crsr.execute(sql_command)
ans = crsr.fetchall()
for i in ans:
    print(i)
connection.close()
