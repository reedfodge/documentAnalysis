import os
import xlrd
import csv
from os import listdir
from os.path import isfile, join
from tabula import convert_into
from zipfile import ZipFile

#Path where spreadsheets are located
path = '/Users/reed/Desktop/testSpreadsheets'
#Gets every file in path and stores in list
files = [f for f in listdir(path) if isfile(join(path, f))]

#Possible file types for files
fileTypes = ['.csv', '.xls', '.xlsx', '.pdf', '.zip']

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
        try:
            wr.writerow(sh.row_values(rownum))
        except UnicodeEncodeError:
            print("UNICODE ENCODE ERROR (FIX)")
    print("CONVERTED XLS(X) FILE " + a + " TO .CSV")
    os.remove(fileName)
    csv_file.close()

#Converts a .pdf file to a .csv
def convertPDF(a):
    fileName = os.path.join(path, a)
    csvFileName = a[0:len(a)-4]
    convert_into(fileName, csvFileName, output_format="csv")
    print("CONVERTED PDF FILE " + a + " TO .CSV")
    os.remove(fileName)

def extractZip(a):
    fileName = os.path.join(path, a)
    with ZipFile(fileName, 'r') as zip:
        zip.extractAll()
    print("EXTRACTED .ZIP FILE " + a)
    os.remove(fileName)

containsZip = False

#Goes through all the files and converts them if necessary
def goThroughFiles():
    for a in files:
        extension = ''
        #Identifies type of file
        for b in fileTypes:
            if(str(a).endswith(b)):
                extension = b
        if(extension == '.xls' or extension == '.xlsx'):
            convertXLS(a)
        elif(extension == '.pdf'):
            convertPDF(a)
        elif(extension == '.zip'):
            extractZip(a)
            containsZip = True

goThroughFiles()
if(containsZip):
    goThroughFiles()
