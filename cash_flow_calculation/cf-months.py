import xlrd
import xlwt
import os
import time
from datetime import datetime
from xlrd import xldate_as_tuple

#input name and dates
excelname = input('File Name (in ".xls" style):')
startdate = int(input('Start Date (yyyymmdd):'))
enddate = int(input('End Date (yyyymmdd):'))

#open workbook and sheets
workbook = xlrd.open_workbook(excelname)
list_of_all_sheetnames = workbook.sheet_names()
total_sheet_num = len(list_of_all_sheetnames)



##total
total = 0
list_of_sum_in_dic = []

#roll through all sheets
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
    for key in dic.keys():
        if key >= startdate and key <= enddate :
            sum_in_dic += dic[key]
    list_of_sum_in_dic.append(sum_in_dic)
    
    total += sum_in_dic
       
#write a new excel
print(list_of_sum_in_dic)
print(total)

wb = xlwt.Workbook(encoding='ascii') 
ws = wb.add_sheet('总计')

ws.write(0, 0, label='公司')
ws.write(0, 1, label='租金总额')

for m in range(0,len(list_of_all_sheetnames)):
    ws.write(m+1, 0, label=list_of_all_sheetnames[m])

for n in range(0,len(list_of_sum_in_dic)):
    ws.write(n+1, 1, label=list_of_sum_in_dic[n])        

ws.write(len(list_of_all_sheetnames)+1, 0, label='总计')
ws.write(len(list_of_all_sheetnames)+1, 1, label=total)



##month
total_per_month = []
month_date_list = []

#month num between the startdate and the enddate
year_total = (enddate // 10000) - (startdate // 10000)
month_num = year_total * 12 + ((enddate % 10000) // 100) - ((startdate % 10000) // 100)

startdate_formonth = startdate
enddate_formonth = enddate

#roll through the months
for k in range(0, month_num):
    sum_in_dic = 0
    
    #roll through all sheets 
    for ii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[ii])
        total_row_num = sheet.nrows
        dic = {}      

        #roll throught the date
        for jj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jj,2).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if (startdate_formonth % 10000) // 100 <= 11:
                if key >= startdate_formonth and key < (startdate_formonth + 100) :
                    sum_in_dic += dic[key]
            else:
                if key >= startdate_formonth and key < (startdate_formonth + 10000 - 1100) :
                    sum_in_dic += dic[key]
        
    total_per_month.append(sum_in_dic)

    month_date_list.append(startdate_formonth)

    if ((startdate_formonth + 100) % 10000) // 100 != 13:
        startdate_formonth += 100
    else:
        startdate_formonth += (10000-1100)

#write a new excel
print(total_per_month)

ws = wb.add_sheet('每月现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')

for mm in range(0,month_num):
    ws.write(mm+1, 0, label=month_date_list[mm])

for nn in range(0,month_num):
    ws.write(nn+1, 1, label=total_per_month[nn])

total_of_all_month = 0
for kk in range(0, month_num):
    total_of_all_month += total_per_month[kk]

ws.write(month_num+1, 0, label='总计')
ws.write(month_num+1, 1, label=total_of_all_month)

newname = '统计数据' + str(startdate) + '-' + str(enddate) + '.xls'
wb.save(newname)

