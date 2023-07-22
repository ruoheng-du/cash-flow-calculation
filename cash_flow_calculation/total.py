import xlrd
import xlwt
import os
import time
from datetime import datetime
from xlrd import xldate_as_tuple

#input name and dates
excelname = input('File Name (in ".xls" style):')
startdate = int(input('Start Date (yyyymmdd):'))
enddate = input('End Date (yyyymmdd) (if no end, please type / ):')

#open workbook and sheets
workbook = xlrd.open_workbook(excelname)
list_of_all_sheetnames = workbook.sheet_names()
total_sheet_num = len(list_of_all_sheetnames)

total = 0
list_of_sum_in_dic = []

#open one specific sheet
for i in range(0, total_sheet_num):
    sheet = workbook.sheet_by_name(list_of_all_sheetnames[i])
    total_row_num = sheet.nrows
    dic = {}
    sum_in_dic = 0

    #roll throught the date
    for j in range(14, total_row_num-1):
        #get the date and make a dictionary
        date_original = xlrd.xldate_as_tuple(sheet.cell(j,1).value,0)
        date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

        dic[date] = sheet.cell(j,2).value

    #find and count the nums in the dictionary by the startdate and the enddate
    if enddate != '/':
        enddate = int(enddate)
        for key in dic.keys():
            if key >= startdate and key <= enddate :
                sum_in_dic += dic[key]
        list_of_sum_in_dic.append(sum_in_dic)
    else:
        for key in dic.keys():
            if key >= startdate:
                sum_in_dic += dic[key]
        list_of_sum_in_dic.append(sum_in_dic)

    total += sum_in_dic
       
#write a new excel
print(list_of_sum_in_dic)
print(total)

wb = xlwt.Workbook(encoding='ascii') 
ws = wb.add_sheet('总计')

for m in range(0,len(list_of_all_sheetnames)):
    ws.write(m, 0, label=list_of_all_sheetnames[m])

for n in range(0,len(list_of_sum_in_dic)):
    ws.write(n, 1, label=list_of_sum_in_dic[n])        

ws.write(len(list_of_all_sheetnames), 0, label='总计')
ws.write(len(list_of_all_sheetnames), 1, label=total)
if enddate != '/':
    newname = '总计' + str(startdate) + '-' + str(enddate) + '.xls'
else:
    newname = '总计' + str(startdate) + '-' + ' ' + '.xls'
wb.save(newname)

