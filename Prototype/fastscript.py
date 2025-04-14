import pandas as pd
from openpyxl import load_workbook

file_name = "" #file name, place the file in the same folder as script
cols = []   #which cols you want to use 

df = pandas.read_excel(file_name, header=0, usecols=cols )

book = load_workbook(file_name)

with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
    writer.book = book
    writer.sheets = {ws.title: ws for ws in book.worksheets}



# future note to use lambda functions for pattern matching in the RFA message text