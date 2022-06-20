import pyodbc 

conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server=mrp1;'
                      'Database=ESIDEMO;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
cursor.execute('SELECT TOP 100000 * FROM ESIDEMO.dbo.ICFPM')

for row in cursor:
    for x in row:
        p = str(x)


            

import openpyxl
from pathlib import Path


xlsx_file = Path('Documents', 'Book1.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file) 



conn1 = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server=mrp1;'
                      'Database=ESIDEMO;'
                      'Trusted_Connection=yes;')

# Read the active sheet:
sheet = wb_obj.active
count = 0
count1 = 0
for row in sheet:
    count = 0
    for cell in row:
        count+= 1
        if count == count1:
            print(cell.value, end = ' ')

            cursor1 = conn1.cursor()
            cursor1.execute('SELECT * FROM ESIDEMO.dbo.POFVP')
            count2 = 0
            holder =''
            for row1 in cursor1:
                count2 = 0
                for x1 in row1:
                    count2 +=1
                    p = x1
                    if count2 == 2:
                        if(x1[0:3] == 'P50'):
                            holder = x1
                    if cell.value == x1:
                        print(holder)
                        
        if cell.value == 'Manufacturer Part Number':
            count1 = count
            
            
