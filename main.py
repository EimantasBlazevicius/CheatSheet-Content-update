# Reading an excel file using Python
import xlrd
import requests

url = 'https://wtdback.qa.bazaarvoice.com/api/'

# Give the location of the file
loc = ("sheet.xls")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

row = 0
# For row 0 and column 0
# print(sheet.cell_value(row, col))
for row in range(sheet.nrows):
    title = sheet.cell_value(row, 0)
    full_q = sheet.cell_value(row, 1)
    answer = sheet.cell_value(row, 2)
    tags_list = sheet.cell_value(row, 3).split(",")
    tags = []
    for tag in tags_list:
        tags.append({"tag": tag})

    myobj = {"title": title,
             "q": full_q,
             "a": answer,
             "suggested_a": "",
             "n": 0,
             "isPublished": True,
             "email": "eimantas.blazevicius@bazaarvoice.com",
             "nickname": "CheatSheet",
             "t": tags,
             "l": []
             }
    x = requests.post(url, json=myobj)
    row += 1
