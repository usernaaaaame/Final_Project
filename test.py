import pandas as pd
import openpyxl


wb = openpyxl.load_workbook('WorkCycle2.xlsx')
ws_names = wb.sheetnames
ws_names_to_show = list()
for i in range(len(ws_names)):
    if (i % 2) == 0:
        ws_names_to_show.append(ws_names[i])
sheet = wb[ws_names_to_show[0]]
print(sheet)
gap=list()
cnt=0
for i in sheet.rows:
    print(i[0].value,end=' ')
    print(i[1].value)
    if (cnt>0):
        gap.append((i[1].value))
    cnt=+1
print(gap)


#find fir ind 테스트
# list1=['o','a',None,'o','o']
# b='a'
# idx=0
# for i in range(len(list1)):
#     if list1[i]==b:
#         idx = i
#         print(idx)
#         break



