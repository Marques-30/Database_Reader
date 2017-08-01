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
level = raw_input("Update, Create a new Table, Delete Table, or Extract a table to excel: ")
Document = raw_input("Enter the name of the excel document: ")
ID = raw_input("Enter Schema name within Database: ")
table=raw_input("Enter Table name within Database: ")

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

if level.lower() == "create":
  with open (Document + '.csv', 'r') as f:
      reader = csv.reader(f)
      columns = next(reader)
      query = """CREATE TABLE {0}.{1} ({2} varchar(255), {3} varchar(255), {4} varchar(255), {5} varchar(255), {6} varchar(255), {7} varchar(255), {8} varchar(255), {9} varchar(255), {10} varchar(255), {11} varchar(255), {12} varchar(255), {13} varchar(255), {14} varchar(255), {15} varchar(255))""".format(LanID, table.lower(), cut[1], cut[3], cut[5], cut[7], cut[9], cut[11], cut[13], cut[15], cut[17], cut[19], cut[21], cut[23], cut[25])
      connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
      print query
      cursor = connection.cursor()
      cursor.execute(query)
      cursor.close()
      print "Table Created"
#############################################################
elif level.lower() == "update":
  with open (Document +'.csv', 'r') as f:
      reader = csv.reader(f)
      columns = next(reader)
      query = """INSERT INTO {0}.{1} ({2}, {3}, {4}, {5}, {6}) VALUES ('{7}', '{8}', '{9}', '{10}', '{11}')""".format(LanID, table.lower(), cut[1], cut[3], cut[5], cut[7], cut[9], cut[11], cut[13], cut[15], cut[17], cut[19])
      query = query.format(','.join(columns), ','.join('?' * len(columns)))
      print query
      connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
      cursor = connection.cursor()
      for data in reader:
          cursor.execute(query)
      connection.commit()
      cursor.close()
      print "Table Updated"
###############################################################
elif level.lower()=="delete":
	query = """DROP TABLE {0}.{1}""".format(LanID.lower(),table.lower())
  	connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
  	cursor = connection.cursor()
  	cursor.execute(query)
  	cursor.close()
  	print "Table Deleted"
elif level.lower()=="extract":
###################################################################
	connection = cx_Oracle.connect(LanID+'/'+password+'@'+Database.upper())
	cursor = connection.cursor()
	with open(Document+".csv", 'w') as out_file:
	        writer=csv.writer(out_file, lineterminator='\n')
	        query = """SELECT * FROM {0}.{1}""".format(ID, table)
	    	cursor.execute(query)
	    	for row in cursor:
		      writer.writerow(row)
	cursor.close()
#################################################################
else:
    print "Invalid command, please try again."
