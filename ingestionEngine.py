import pandas
import fileConverter
import os

#Path where spreadsheets are located
path = '/Users/reed/Desktop/testSpreadsheets'
#Gets every file in path and stores in list
files = [f for f in listdir(path) if isfile(join(path, f))]

for a in files:
    if(a.endswith('.csv')):
        df = pandas.read_csv(a)
        print(df)
