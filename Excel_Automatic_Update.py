import win32com.client 
import time
import subprocess

#Check if pywin32 module is installed
ps1 = subprocess.call(r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe pip install pywin32' , shell=True)

# Start an instance of Excel
xlapp = win32com.client.DispatchEx("Excel.Application")

#Sheet Range
p1 = 'sheet1'
p2 = 'sheet2'
p3 = 'sheet3'
p4 = 'sheet4'

# Optional, e.g. if you want to debug
xlapp.Visible = True

# Open the workbook in said instance of Excel
wb = xlapp.Workbooks.open(r'C:\File\Path\Excel.xlsb')

#Go through all the pages and update them
for sheet in p1,p2,p3,p4:
	sheet = wb.Worksheets(sheet).Select()
	wb.RefreshAll()
	time.sleep(5)

#Add short delay for debugging
time.sleep(5)
wb.Save()

# Quit the excel
xlapp.Quit()