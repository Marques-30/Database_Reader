import os
import sys
import cx_Oracle
import csv
import subprocess
from xlrd import open_workbook
import xlrd
import getpass

LanID=raw_input("Schema / Username: ")
Database=raw_input("Enter Orcale Database: ")
password=getpass.getpass("Enter Password: ")
Document = raw_input("Enter the name of the excel document: ")
ID = raw_input("Enter Schema name within Database: ")
table=raw_input("Enter Table name within Database: ")
table2=raw_input("Enter new Table name within Database: ")


wb = open_workbook(Document+'.xlsx')
for s in wb.sheets():
    print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append(value)
        values.append(col_value)
print values

def csv_from_excel():

	wb = xlrd.open_workbook(Document+'.xlsx')
	sh = wb.sheet_by_name('Sheet1')
	Report = open(Document + '.csv', 'wb')
	wr = csv.writer(Report, quoting=csv.QUOTE_ALL)

	for rownum in xrange(sh.nrows):
	    wr.writerow(sh.row_values(rownum))

	Report.close()

cut=str(values).split("'")

if __name__ == '__main__':
  	csv_from_excel()

connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
cursor = connection.cursor()
with open(Document+".csv", 'w') as out_file:
        writer=csv.writer(out_file, lineterminator='\n')
        query = """SELECT * FROM {0}.{1}""".format(ID, table)
        cursor.execute(query)
        col_names = []
        for i in range(0, len(cursor.description)):
            col_names.append(cursor.description[i][0])
            data_col = str(col_names).split("'")
            print data_col[1], data_col[3], data_col[5], data_col[7], data_col[9]
        for row in cursor:
          writer.writerow(row)
cursor.close()


with open (Document + '.csv', 'r') as f:
    connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
    cursor = connection.cursor()
    reader = csv.reader(f)
    create = """CREATE TABLE {0}.{1} ({2} varchar(255), {3} varchar(255), {4} varchar(255), {5} varchar(255), {6} varchar(255))""".format(LanID, table2.lower(), cut[1], cut[3], cut[5], cut[7], cut[9])
    print create
    cursor.execute(create)
    print "Table Created"
    count = 0
    while count < 150:
      query = """INSERT INTO {0}.{1} ({2}, {3}, {4}, {5}, {6}) VALUES ('{7}', '{8}', '{9}', '{10}', '{11}')""".format(LanID, table2.lower(), cut[1], cut[3], cut[5], cut[7], cut[9], cut[11 + count], cut[13 + count], cut[15 + count], cut[17 + count], cut[19 + count])
      print query
      cursor.execute(query)
      connection.commit()
      count += 10
    cursor.close()
    print "Table Updated"
  