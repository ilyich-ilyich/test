from openpyxl import load_workbook
import pandas as pd


df = pd.read_excel('c:/bu/1.xlsx')



wb = load_workbook('c:/bu/2.xlsx')
ws = wb.active
#print(wb)
print(ws.cell(row=6, column=9).hyperlink.target)
#print(ws.cell(row=5, column=8).value)



links = []
#for i in range(2, ws.max_row + 1):  # 2nd arg in range() not inclusive, so add 1
    #links.append(ws.cell(row=i, column=1).hyperlink.target)

#df['link'] = links