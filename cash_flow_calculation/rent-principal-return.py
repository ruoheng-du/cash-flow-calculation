import xlrd
import xlwt

print('注意：1.表格格式为xls 2.租期第一期从第15行开始 3.日期在B列 4.租金在C列 5.每期本金在F列 6.每期利息在G列')

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
total_princ = 0
total_interest = 0
list_of_sum_in_dic = []
list_of_sum_in_dic_princ = []
list_of_sum_in_dic_interest = []

#roll through all sheets
for i in range(0, total_sheet_num):
    sheet = workbook.sheet_by_name(list_of_all_sheetnames[i])
    total_row_num = sheet.nrows
    dic = {}
    dic_princ = {}
    dic_interest = {}
    sum_in_dic = 0
    sum_in_dic_princ = 0
    sum_in_dic_interest = 0

    #roll throught the date
    for j in range(14, total_row_num-1):
        #get the date and make a dictionary
        date_original = xlrd.xldate_as_tuple(sheet.cell(j,1).value,0)
        date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

        dic[date] = sheet.cell(j,2).value
        dic_princ[date] = sheet.cell(j,5).value
        dic_interest[date] = sheet.cell(j,6).value  

    #find and count the nums in the dictionary by the startdate and the enddate
    for key in dic.keys():
        if key >= startdate and key <= enddate :
            sum_in_dic += dic[key]
            sum_in_dic_princ += dic_princ[key]
            sum_in_dic_interest += dic_interest[key]
    list_of_sum_in_dic.append(sum_in_dic)
    list_of_sum_in_dic_princ.append(sum_in_dic_princ)
    list_of_sum_in_dic_interest.append(sum_in_dic_interest)
    
    total += sum_in_dic
    total_princ += sum_in_dic_princ
    total_interest =+ sum_in_dic_interest
       
#write a new excel
print(list_of_sum_in_dic)
print(list_of_sum_in_dic_princ)
print(list_of_sum_in_dic_interest)
print(total)

wb = xlwt.Workbook(encoding='ascii') 
ws = wb.add_sheet('总计')

ws.write(0, 0, label='公司')
ws.write(0, 1, label='租金总额')
ws.write(0, 2, label='本金总额')
ws.write(0, 3, label='收益总额')

for m in range(0,len(list_of_all_sheetnames)):
    ws.write(m+1, 0, label=list_of_all_sheetnames[m])

for n in range(0,len(list_of_sum_in_dic)):
    ws.write(n+1, 1, label=list_of_sum_in_dic[n])

for f in range(0, len(list_of_sum_in_dic_princ)):
    ws.write(f+1, 2, label=list_of_sum_in_dic_princ[f])

for g in range(0, len(list_of_sum_in_dic_interest)):
    ws.write(g+1, 3, label=list_of_sum_in_dic_interest[g])

ws.write(len(list_of_all_sheetnames)+1, 0, label='总计')
ws.write(len(list_of_all_sheetnames)+1, 1, label=total)
ws.write(len(list_of_all_sheetnames)+1, 2, label=total_princ)
ws.write(len(list_of_all_sheetnames)+1, 3, label=total_interest)



##month
total_per_month = []
total_per_month_princ = []
total_per_month_interest = []
month_date_list = []

#month num between the startdate and the enddate
year_total = (enddate // 10000) - (startdate // 10000)
month_num = year_total * 12 + ((enddate % 10000) // 100) - ((startdate % 10000) // 100)

startdate_formonth = startdate
enddate_formonth = enddate

#roll through the months
for k in range(0, month_num):
    sum_in_dic = 0
    sum_in_dic_princ = 0
    sum_in_dic_interest = 0
    
    #roll through all sheets 
    for ii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[ii])
        total_row_num = sheet.nrows
        dic = {}
        dic_princ = {}
        dic_interest = {}

        #roll throught the date
        for jj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jj,2).value
            dic_princ[date] = sheet.cell(jj,5).value
            dic_interest[date] = sheet.cell(jj,6).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if (startdate_formonth % 10000) // 100 <= 11:
                if key >= startdate_formonth and key < (startdate_formonth + 100) :
                    sum_in_dic += dic[key]
                    sum_in_dic_princ += dic_princ[key]
                    sum_in_dic_interest += dic_interest[key]
            else:
                if key >= startdate_formonth and key < (startdate_formonth + 10000 - 1100) :
                    sum_in_dic += dic[key]
                    sum_in_dic_princ += dic_princ[key]
                    sum_in_dic_interest += dic_interest[key]
        
    total_per_month.append(sum_in_dic)
    total_per_month_princ.append(sum_in_dic_princ)
    total_per_month_interest.append(sum_in_dic_interest)

    month_date_list.append(startdate_formonth)

    if ((startdate_formonth + 100) % 10000) // 100 != 13:
        startdate_formonth += 100
    else:
        startdate_formonth += (10000-1100)

#write a new excel
print(total_per_month)
print(total_per_month_princ)
print(total_per_month_interest)

ws = wb.add_sheet('每月现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')
ws.write(0, 2, label='本金总额')
ws.write(0, 3, label='收益总额')

for mm in range(0,month_num):
    ws.write(mm+1, 0, label=month_date_list[mm])

for nn in range(0,month_num):
    ws.write(nn+1, 1, label=total_per_month[nn])

for ff in range(0,month_num):
    ws.write(ff+1, 2, label=total_per_month_princ[ff])

for gg in range(0,month_num):
    ws.write(gg+1, 3, label=total_per_month_interest[gg])

total_of_all_month = 0
for kk in range(0, month_num):
    total_of_all_month += total_per_month[kk]

total_of_all_month_princ = 0
for qq in range(0, month_num):
    total_of_all_month_princ += total_per_month_princ[qq]

total_of_all_month_interest = 0
for pp in range(0, month_num):
    total_of_all_month_interest += total_per_month_interest[pp]

ws.write(month_num+1, 0, label='总计')
ws.write(month_num+1, 1, label=total_of_all_month)
ws.write(month_num+1, 2, label=total_of_all_month_princ)
ws.write(month_num+1, 3, label=total_of_all_month_interest)



##season
total_per_season = []
total_per_season_princ = []
total_per_season_interest = []
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
    sum_in_dic_princ = 0
    sum_in_dic_interest = 0
    
    #roll through all sheets 
    for iii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[iii])
        total_row_num = sheet.nrows
        dic = {}
        dic_princ = {}
        dic_interest = {}

        #roll throught the date
        for jjj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jjj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jjj,2).value
            dic_princ[date] = sheet.cell(jjj,5).value
            dic_interest[date] = sheet.cell(jjj,6).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if (startdate_forseason % 10000) // 100 <= 9:
                if key >= startdate_forseason and key < (startdate_forseason + 300) :
                    sum_in_dic += dic[key]
                    sum_in_dic_princ += dic_princ[key]
                    sum_in_dic_interest += dic_interest[key]
            else:
                if key >= startdate_forseason and key < (startdate_forseason + 10000 - 900) :
                    sum_in_dic += dic[key]
                    sum_in_dic_interest += dic_interest[key]
        
    total_per_season.append(sum_in_dic)
    total_per_season_princ.append(sum_in_dic_princ)
    total_per_season_interest.append(sum_in_dic_interest)

    season_date_list.append(startdate_forseason)

    if (startdate_forseason % 10000) // 100 <= 9:
        startdate_forseason += 300
    else:
        startdate_forseason += (10000-900)

#write a new excel
print(total_per_season)
print(total_per_season_princ)
print(total_per_season_interest)

ws = wb.add_sheet('每季现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')
ws.write(0, 2, label='本金总额')
ws.write(0, 3, label='收益总额')

for mmm in range(0,season_num):
    ws.write(mmm+1, 0, label=season_date_list[mmm])

for nnn in range(0,season_num):
    ws.write(nnn+1, 1, label=total_per_season[nnn])

for fff in range(0,season_num):
    ws.write(fff+1, 2, label=total_per_season_princ[fff])

for ggg in range(0,season_num):
    ws.write(ggg+1, 3, label=total_per_season_interest[ggg])

total_of_all_season = 0
for uuu in range(0, season_num):
    total_of_all_season += total_per_season[uuu]

total_of_all_season_princ = 0
for qqq in range(0, season_num):
    total_of_all_season_princ += total_per_season_princ[qqq]

total_of_all_season_interest = 0
for ppp in range(0, season_num):
    total_of_all_season_interest += total_per_season_interest[ppp]

ws.write(season_num+1, 0, label='总计')
ws.write(season_num+1, 1, label=total_of_all_season)
ws.write(season_num+1, 2, label=total_of_all_season_princ)
ws.write(season_num+1, 3, label=total_of_all_season_interest)



##natural quarter
total_per_natural = []
total_per_natural_princ = []
total_per_natural_interest = []
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
    sum_in_dic_princ = 0
    sum_in_dic_interest = 0
    
    #roll through all sheets 
    for iiii in range(0, total_sheet_num):
        sheet = workbook.sheet_by_name(list_of_all_sheetnames[iiii])
        total_row_num = sheet.nrows
        dic = {}
        dic_princ = {}
        dic_interest = {}

        #roll throught the date
        for jjjj in range(14, total_row_num-1):
            #get the date and make a dictionary
            date_original = xlrd.xldate_as_tuple(sheet.cell(jjjj,1).value,0)
            date = date_original[0]*10000 + date_original[1]*100 + date_original[2]*1

            dic[date] = sheet.cell(jjjj,2).value
            dic_princ[date] = sheet.cell(jjjj,5).value
            dic_interest[date] = sheet.cell(jjjj,6).value

        #find and count the nums in the dictionary by the startdate and the enddate
        for key in dic.keys():
            if kkkk != len(natural_date_list)-1:
                if key >= natural_date_list[kkkk] and key <= natural_date_list[kkkk+1]:
                    sum_in_dic += dic[key]
                    sum_in_dic_princ += dic_princ[key]
                    sum_in_dic_interest += dic_interest[key]
            else:
                if key >= natural_date_list[kkkk] and key <= enddate_fornatural:
                    sum_in_dic += dic[key]
                    sum_in_dic_princ += dic_princ[key]
                    sum_in_dic_interest += dic_interest[key]
        
    total_per_natural.append(sum_in_dic)
    total_per_natural_princ.append(sum_in_dic_princ)
    total_per_natural_interest.append(sum_in_dic_interest)

#write a new excel
print(total_per_natural)
print(total_per_natural_princ)
print(total_per_natural_interest)

ws = wb.add_sheet('自然季度现金流')

ws.write(0, 0, label='起始日')
ws.write(0, 1, label='租金总额')
ws.write(0, 2, label='本金总额')
ws.write(0, 3, label='收益总额')

for mmmm in range(0,len(natural_date_list)):
    ws.write(mmmm+1, 0, label=natural_date_list[mmmm])

for nnnn in range(0,len(natural_date_list)):
    ws.write(nnnn+1, 1, label=total_per_natural[nnnn])

for ffff in range(0,len(natural_date_list)):
    ws.write(ffff+1, 2, label=total_per_natural_princ[ffff])

for gggg in range(0,len(natural_date_list)):
    ws.write(gggg+1, 3, label=total_per_natural_interest[gggg])

total_of_all_natural = 0
for uuuu in range(0, len(natural_date_list)):
    total_of_all_natural += total_per_natural[uuuu]

total_of_all_natural_princ = 0
for qqqq in range(0, len(natural_date_list)):
    total_of_all_natural_princ += total_per_natural_princ[qqqq]

total_of_all_natural_interest = 0
for pppp in range(0, len(natural_date_list)):
    total_of_all_natural_interest += total_per_natural_interest[pppp]

ws.write(len(natural_date_list)+1, 0, label='总计')
ws.write(len(natural_date_list)+1, 1, label=total_of_all_natural)
ws.write(len(natural_date_list)+1, 2, label=total_of_all_natural_princ)
ws.write(len(natural_date_list)+1, 3, label=total_of_all_natural_interest)


#form a new excel
newname = '统计数据' + str(startdate) + '-' + str(enddate) + '.xls'
wb.save(newname)

