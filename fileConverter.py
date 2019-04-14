import os
from os import listdir
from os.path import isfile, join
import xlrd
import csv

#Path where spreadsheets are located
path = '/Users/reed/Desktop/testSpreadsheets'
#Gets every file in path and stores in list
files = [f for f in listdir(path) if isfile(join(path, f))]

#Possible file types for files
fileTypes = ['.csv', '.xls', '.xlsx', '.pdf']

#Converts a .xls or .xlsx file to a .csv file
def convertXLS(a):
    nameWithoutExtension = ''
    if(a.endswith('.xls')):
        nameWithoutExtension = a[0:len(a)-4]
    else:
        nameWithoutExtension = a[0:len(a)-5]
    fileName = os.path.join(path, a)
    wb = xlrd.open_workbook(fileName)
    sh = wb.sheet_by_index(0)
    nameWithoutExtension += '.csv'
    csv_file = open(nameWithoutExtension, 'wb')
    wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))
    print("CONVERTED FILE " + a + " TO .CSV")
    csv_file.close()


#Converts a .pdf file to a .csv
#def convertPDF(a):


for a in files:
    extension = ''
    #Identifies type of file
    for b in fileTypes:
        if(str(a).endswith(b)):
            extension = b
    if(extension == '.xls' or extension == '.xlsx'):
        convertXLS(a)
    #elif(extension == '.pdf'):
        #convertPDF(a)
