import pandas as pd
import win32com.client as win32
import os

path = os.path.join(os.getcwd(), "Templates/Template.xlsx")

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(path)
ws = wb.Worksheets("Sheet1")
excel.Visible = True

current_row = 1
current_column = 1

data_dict = {
  1:{
    "Name": "John",
    "Address": "221 Icon Street",
    "Phone Number": "8293032922",
    "Product": "Apples"
  }
}

while ws.Cells(current_row, current_column).Value:
  current_row += 1
else:
  for data in data_dict.values():
    ws.Cells(current_row, current_column).Value = data["Name"]
    ws.Cells(current_row, current_column).Offset(1,2).Value = data["Address"]
    ws.Cells(current_row, current_column).Offset(1,3).Value = data["Phone Number"]
    ws.Cells(current_row, current_column).Offset(1,4).Value = data["Product"]
    
    current_row += 1
  
  print("No more data to process")
