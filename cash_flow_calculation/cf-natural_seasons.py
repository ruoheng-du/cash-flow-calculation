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



##season
total_per_season = []
season_date_list = []

#season num between the startdate and the enddate
year_total = (enddate // 10000) - (startdate // 10000)
month_total = year_total * 12 + ((enddate % 10000) // 100) - ((startdate % 10000) // 100)
if month_total % 3 == 0:
    season_num = int(month_total / 3)
else:
    season_num = int(month_total // 3 + 1)

startdate_forseason = startdate
eneddate_forseason = enddate

#roll through the seasons
for kkk in range(0, season_num):
    sum_in_dic = 0
    
    #roll through all sheets 
    for iii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[iii])
        total_row_num = sheet.nrows
        dic = {}      

        #roll throught the date
        for jjj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jjj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jjj,2).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if (startdate_forseason % 10000) // 100 <= 9:
                if key >= startdate_forseason and key < (startdate_forseason + 300) :
                    sum_in_dic += dic[key]
            else:
                if key >= startdate_forseason and key < (startdate_forseason + 10000 - 900) :
                    sum_in_dic += dic[key]
        
    total_per_season.append(sum_in_dic)

    season_date_list.append(startdate_forseason)

    if (startdate_forseason % 10000) // 100 <= 9:
        startdate_forseason += 300
    else:
        startdate_forseason += (10000-900)

#write a new excel
print(total_per_season)

ws = wb.add_sheet('每季现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')

for mmm in range(0,season_num):
    ws.write(mmm+1, 0, label=season_date_list[mmm])

for nnn in range(0,season_num):
    ws.write(nnn+1, 1, label=total_per_season[nnn])

total_of_all_season = 0
for uuu in range(0, season_num):
    total_of_all_season += total_per_season[uuu]

ws.write(season_num+1, 0, label='总计')
ws.write(season_num+1, 1, label=total_of_all_season)



##natural quarter
total_per_natural = []
natural_date_list = []

startdate_fornatural = startdate
enddate_fornatural = enddate

#natural quarter num between the startdate and the enddate
year_total = (enddate // 10000) - (startdate // 10000)
month_total = year_total * 12 + ((enddate % 10000) // 100) - ((startdate % 10000) // 100)

natural_date_list.append(startdate)

#find the first natural date
if (startdate_fornatural % 10000) < 401:
    first_natural_date = (startdate_fornatural // 10000) * 10000 + 401
elif (startdate_fornatural % 10000) < 701:
    first_natural_date = (startdate_fornatural // 10000) * 10000 + 701
elif (startdate_fornatural % 10000) < 1001:
    first_natural_date = (startdate_fornatural // 10000) * 10000 + 1001
else:
    first_natural_date = (startdate_fornatural // 10000) * 10000 + 10000 + 101
if first_natural_date < enddate_fornatural:
    natural_date_list.append(first_natural_date)

natural_date = first_natural_date
while True:
    if (natural_date % 10000) != 1001:
        natural_date += 300
    else:
        natural_date += (10000-900)
    if natural_date >= enddate_fornatural:
        break
    natural_date_list.append(natural_date)

#roll through the seasons
for kkkk in range(0, len(natural_date_list)):
    sum_in_dic = 0
    
    #roll through all sheets 
    for iiii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[iiii])
        total_row_num = sheet.nrows
        dic = {}      

        #roll throught the date
        for jjjj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jjjj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jjjj,2).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if kkkk != len(natural_date_list)-1:
                if key >= natural_date_list[kkkk] and key <= natural_date_list[kkkk+1]:
                    sum_in_dic += dic[key]
            else:
                if key >= natural_date_list[kkkk] and key <= enddate_fornatural:
                    sum_in_dic += dic[key]
        
    total_per_natural.append(sum_in_dic)

#write a new excel
print(total_per_natural)

ws = wb.add_sheet('自然季度现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')

for mmmm in range(0,len(natural_date_list)):
    ws.write(mmmm+1, 0, label=natural_date_list[mmmm])

for nnnn in range(0,len(natural_date_list)):
    ws.write(nnnn+1, 1, label=total_per_natural[nnnn])

total_of_all_natural = 0
for uuuu in range(0, len(natural_date_list)):
    total_of_all_natural += total_per_natural[uuuu]

ws.write(len(natural_date_list)+1, 0, label='总计')
ws.write(len(natural_date_list)+1, 1, label=total_of_all_natural)

newname = '统计数据' + str(startdate) + '-' + str(enddate) + '.xls'
wb.save(newname)

