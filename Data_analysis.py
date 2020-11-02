import sys
import win32com.client

xlApp = win32com.client.Dispatch("Excel.Application")
print("Excel library version:", xlApp.Version)
filename,password = 'C:\myfiles\foo.xls', 'qwerty12'
xlwb = xlApp.Workbooks.Open(filename, Password=password)