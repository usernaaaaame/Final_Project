import pandas as pd
import openpyxl


wb = openpyxl.load_workbook('sample.xlsx')
ws_names = wb.sheetnames
ws_names_to_show = list()
for i in range(len(ws_names)):
    if (i % 2) == 0:
        ws_names_to_show.append(ws_names[i])
sheet = wb[ws_names_to_show[1]]
print(sheet)
gap=list()
for i in sheet.columns:
    gap.append(i[3].value)
print(gap)
wb.create_sheet('testsheet')
wb.save('sample.xlsx')

# # 열 출력 예시
# cnt=0
# for i in sheet.rows:
#     print(i[0].value,end=' ')
#     print(i[1].value)
#     if (cnt>0):
#         gap.append((i[1].value))
#     cnt=+1
# del gap[-1]

#find fir ind 테스트
# result=list()
# list1=['o','a',None,'a','a']
# b='a'
# for i in range(len(list1)):
#     if list1[i]==b:
#         result.append(i)
#         print(i)
# print(result)

# find idx Lists 테스트
# list1 = ["가","나","다","나"]
# list2 = ["o",None,"o","o"]
# val1="라"
# val2="X"
# for i in range(len(list1)):
#     if list1[i]==val1 and list2[i]==val2 :
#         print(i)
#
# idx_list=list()
# for i in range(len(list1)):
#     if list1[i]==val1:
#         idx_list.append(i)
# print(idx_list)
# print(len(idx_list))