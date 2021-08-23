import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import string


#Pandas to read the Excel file, create a pivot table, and export it to Excel
#Then we’ll use the Openpyxl library to write Excel formulas, make charts and format the spreadsheet through Python

excel_file = pd.read_excel('Impulse Noise Evaluation.xlsx')
excel_file[[]]
report_table = excel_file.pivot_table(index='Measurement time')
report_table.to_excel('report_20212323.xlsx', sheet_name='Report', startrow=4)#Creates report in excel

wb = load_workbook('report_20212323.xlsx')# ariable to load workbook
sheet = wb['Report']# assign sheet

# cell references (original spreadsheet)
min_column = wb.active.min_column
max_column = wb.active.max_column
min_row = wb.active.min_row
max_row = wb.active.max_row




print(report_table)
#print(excel_file)
