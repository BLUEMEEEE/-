import xlrd
import xlwt
import datetime
import re


# 定位
def search_for_word(sheet,word):
    index=-1
    nrows=sheet.nrows
    for i in range(nrows):
        if(i==0):
            i=i+1
        row=sheet.row_values(i)
        if(re.search(word,str(row[0]))!=None):
            index=i
            break
    return index

# 研究生不统计
def is_undergraduate(sheet,row_num):
    row=sheet.row_values(row_num)
    if(row[9]=="本科生"):
        return True
    else:
        return False

# 查看是否和上一条数据重复(可能出现该人又刷通道又刷门，此时会有两条并列的数据)
def is_duplicate(sheet,row_num):
    row=sheet.row_values(row_num)
    if(row_num-1<0):
        return False
    else:
        last_row=sheet.row_values(row_num-1)
        if(row[1]==last_row[1]):
            return True
        else:
            return False

# 将row 里面的内容加入到output_sheet里面去
def write_in_output_sheet(output_sheet,output_line_counting,sheet,row_num):
    row=sheet.row_values(row_num)
    for i in range(len(row)):
        if(i==7):
            if(sheet.cell(row_num,i).ctype==3):
                date = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(row_num,i).value,0))
                cell = date.strftime('%Y/%m/%d')
            else:
                cell=sheet.cell(row_num,i).value
            output_sheet.write(output_line_counting,i,label=cell)
        elif(i==8):
            if(sheet.cell(row_num,i).ctype==3):
                tup=xlrd.xldate_as_tuple(sheet.cell(row_num,i).value,0)
                cell=str(tup[-3])+":"+str(tup[-2])+":"+str(tup[-1])
            else:
                cell=sheet.cell(row_num,i).value
            output_sheet.write(output_line_counting,i,label=cell)
        #if(i==8):
        #    tup=xlrd.xldate_as_tuple(sheet.cell(2,8).value,0)
        #    tup[0]=1999
        #    tup[1]=9
        #    tup[2]=9
        #    time=datetime.datetime(*tup)
        #    style = xlwt.XFStyle()
        #    style.num_format_str = 'M/D/YY'
        #    output_sheet.write(output_line_counting, i, time, style)
        #elif(i==9):
        #    time=xlrd.xldate.xldate_as_datetime(sheet.cell(row_num,i).value,0)
        #    style = xlwt.XFStyle()
        #    style.num_format_str = 'h:mm:ss'
        #    output_sheet.write(output_line_counting, i, time, style)
        else:
            output_sheet.write(output_line_counting,i,label=row[i])
    output_line_counting=output_line_counting+1
    return output_line_counting

# 是否晚于11:30回宿舍
def is_later(sheet,row_num):
    if(sheet.cell(row_num,8).ctype==3):
        tup=xlrd.xldate_as_tuple(sheet.cell(row_num,8).value,0)
        cell=str(tup[-3])+":"+str(tup[-2])+":"+str(tup[-1])
    else:
        tup=sheet.cell(row_num,8).value.split(":")
        cell=sheet.cell(row_num,8).value
    if(int(tup[-3])<=7):
        return True
    elif(int(tup[-3])==23):
        if(int(tup[-2])>=30):
            return True
        else:
            return False
    else:
        return False

def search_for_return_record(sheet,part2_index,row_num):
    if(sheet.cell(row_num,7).ctype==3):
        temp=xlrd.xldate_as_tuple(sheet.cell(int(row_num),7).value,0)
        date = datetime.datetime(*temp).strftime('%Y/%m/%d')
    else:
        date=sheet.cell(int(row_num),7).value
    for i in range(sheet.nrows):
        if(i<=row_num):
            i=row_num
        row=sheet.row_values(i)
        if(row[0]==''):
            continue
        if(row[0]=='学号'):
            continue
        if(not is_undergraduate(sheet,i)):# 研究生跳过
            continue
        if(row[1]==sheet.cell(row_num,1).value):
            if(row[6]=="进入宿舍"):
                if(sheet.cell(i,7).ctype==3):
                    temp = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(i,7).value,0)).strftime('%Y/%m/%d')
                else:
                    temp=sheet.cell(i,7).value
                if(date==temp):
                    return True
    return False

# 打开文件
path="D:\\MyFiles\\思教科\\2018年下半年晚归数据\\"
file_names=["林主任-第一周各学院晚归记录报表（2018.09.02-2018.09.08)",
"林主任-第二周各学院晚归记录报表（2018.09.09-2018.09.15)",
"林主任-第三周各学院晚归记录报表（2018.09.16-2018.09.22)",
"林主任-第四周各学院晚归记录报表(2018.09.22-2018.09.30)",
"林主任-第六周各学院晚归记录报表(2018.10.07-2018.10.13)",
"林主任-第七周各学院晚归记录报表(2018.10.14-2018.10.20)",
"林主任-第八周各学院晚归记录报表(2018.10.21-2018.10.27)",
"林主任-第九周各学院晚归记录报表(2018.10.28-2018.11.3)",
"林主任-第十周各学院晚归记录报表(2018.11.4-2018.11.10)",
"林主任-第十一周各学院晚归记录报表(2018.11.11-2018.11.17)",
"林主任-第十二周各学院晚归记录报表(2018.11.18-2018.11.24)",
"林主任-第十三周各学院晚归记录（11月25日—12月2日）",
"林主任-第十四周各学院晚归记录（12月2日—12月9日）",
"林主任-第15周晚归情况总表1209_1216",
"林主任-第十六周学院总表1216_1223",
"林主任-第十七周学院总表1223_1230",
"林主任-第十八周学院总表1230_0106（新）",
"林主任-第十九周学院总表0106_0113"]
output_base_path="D:\\MyFiles\\思教科\\2018年下半年晚归数据\\output\\"
output_type=['所有晚归\\','11.30后晚归\\','晚归后离开未回宿舍\\','晚归后离开宿舍后再回\\']

# 计算所有晚归数据
'''
for file_index in range(len(file_names)):
    if(file_index==12):# 发现十四周的数据是空的，跳过
        continue
    workbook = xlrd.open_workbook(path+file_names[file_index]+".xlsx")
    print("*******打开"+file_names[file_index]+".xlsx"+"*******")

    sheet_names=workbook.sheet_names()

    output_workbook_1=xlwt.Workbook(encoding="utf-8")# new一个输出的output_
    for m in range(len(sheet_names)):
        sheet_name=sheet_names[m]
        if(sheet_name=="汇总表"):
            continue
        print(sheet_names[m])
        sheet=workbook.sheet_by_name(sheet_name)# 找到以sheet_name为名字的sheet
        output_line_counting=0
        output_sheet=output_workbook_1.add_sheet(sheet_name)# 增加一个sheet
        part2_index=search_for_word(sheet,"晚归")# 定位"晚归时间后离开宿舍"的位置
        nrows=sheet.nrows# 行数
        ncols=sheet.ncols# 列数
        
        for i in range(nrows):
            if(i>=part2_index):
                break
            row=sheet.row_values(i)
            if(row[0]==''):
                continue
            if(row[0]=='学号'):
                continue
            if(row[1]==''):
                continue
            if(row[0]=='晚归时间后进入宿舍'):
                continue
            if(not is_undergraduate(sheet,i)):# 研究生跳过
                continue
            if(is_duplicate(sheet,i)):# 如果重复了跳过
                continue
            output_line_counting=write_in_output_sheet(output_sheet,output_line_counting,sheet,i)
    output_workbook_1.save(output_base_path+output_type[0]+file_names[file_index]+".xls")
'''
'''
# 计算23：30后晚归

for file_index in range(len(file_names)):
    if(file_index==12):# 发现十四周的数据是空的，跳过
        continue
    workbook = xlrd.open_workbook(path+file_names[file_index]+".xlsx")
    print("*******打开"+file_names[file_index]+".xlsx"+"*******")

    sheet_names=workbook.sheet_names()

    output_workbook=xlwt.Workbook(encoding="utf-8")# new一个输出的output_
    for m in range(len(sheet_names)):
        sheet_name=sheet_names[m]
        if(sheet_name=="汇总表"):
            continue
        print(sheet_names[m])
        sheet=workbook.sheet_by_name(sheet_name)# 找到以sheet_name为名字的sheet
        output_line_counting=0
        output_sheet=output_workbook.add_sheet(sheet_name)# 增加一个sheet
        part2_index=search_for_word(sheet,"晚归")# 定位"晚归时间后离开宿舍"的位置
        nrows=sheet.nrows# 行数
        ncols=sheet.ncols# 列数
        
        for i in range(nrows):
            if(i>=part2_index):
                break
            row=sheet.row_values(i)
            if(row[0]==''):
                continue
            if(row[0]=='学号'):
                continue
            if(row[1]==''):
                continue
            if(row[0]=='晚归时间后进入宿舍'):
                continue
            if(not is_undergraduate(sheet,i)):# 研究生跳过
                continue
            if(is_duplicate(sheet,i)):# 如果重复了跳过
                continue
            if(not is_later(sheet,i)):# 如果不是太晚的跳过
                continue
            output_line_counting=write_in_output_sheet(output_sheet,output_line_counting,sheet,i)
    output_workbook.save(output_base_path+output_type[1]+file_names[file_index]+".xls")
'''

# 晚归时间离开宿舍未回宿舍
# 晚归时间离开宿舍后回寝
for file_index in range(len(file_names)):
    if(file_index==12):# 发现十四周的数据是空的，跳过
        continue
    workbook = xlrd.open_workbook(path+file_names[file_index]+".xlsx")
    print("*******打开"+file_names[file_index]+".xlsx"+"*******")

    sheet_names=workbook.sheet_names()

    output_workbook_1=xlwt.Workbook(encoding="utf-8")# 晚归时间离开宿舍后回寝
    output_workbook_2=xlwt.Workbook(encoding="utf-8")# 晚归时间离开宿舍未回宿舍
    for m in range(len(sheet_names)):
        sheet_name=sheet_names[m]
        if(sheet_name=="汇总表"):
            continue
        print(sheet_names[m])
        sheet=workbook.sheet_by_name(sheet_name)# 找到以sheet_name为名字的sheet
        output_line_counting_1=0# 晚归时间离开宿舍后回寝
        output_line_counting_2=0# 晚归时间离开宿舍未回宿舍
        output_sheet_1=output_workbook_1.add_sheet(sheet_name)# 增加一个sheet
        output_sheet_2=output_workbook_2.add_sheet(sheet_name)# 增加一个sheet
        part2_index=search_for_word(sheet,"晚归")# 定位"晚归时间后离开宿舍"的位置
        nrows=sheet.nrows# 行数
        ncols=sheet.ncols# 列数
        
        for i in range(nrows):
            if(i<part2_index):
                i=part2_index
            row=sheet.row_values(i)
            if(row[0]==''):
                continue
            if(row[0]=='学号'):
                continue
            if(row[0]=='晚归时间后进入宿舍'):
                continue
            if(row[1]==''):
                continue
            if(not is_undergraduate(sheet,i)):# 研究生跳过
                continue
            if(is_duplicate(sheet,i)):# 如果重复了跳过
                continue
            if(row[6]=="离开宿舍"):
                if(search_for_return_record(sheet,part2_index,i)):
                    output_line_counting_1=write_in_output_sheet(output_sheet_1,output_line_counting_1,sheet,i)
                else:
                    output_line_counting_2=write_in_output_sheet(output_sheet_2,output_line_counting_2,sheet,i)
    output_workbook_1.save(output_base_path+output_type[3]+file_names[file_index]+".xls")
    output_workbook_2.save(output_base_path+output_type[2]+file_names[file_index]+".xls")


