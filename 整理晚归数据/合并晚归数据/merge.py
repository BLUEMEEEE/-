import os
import xlrd
import xlwt
import datetime
import re

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
        else:
            output_sheet.write(output_line_counting,i,label=row[i])
    output_line_counting=output_line_counting+1
    return output_line_counting

base_path=os.getcwd()
input_path=base_path+"\\input"
output_path=base_path+"\\output"
print("base_path:"+base_path)

for root,dirs,files in os.walk(input_path):
    print("files for merging:")
    print(files)
    workbooks=[]
    for fileName in files:
        print(input_path+"\\"+fileName)
        workbooks.append(xlrd.open_workbook(input_path+"\\"+fileName))
    sheet_names=workbooks[0].sheet_names()
    output_workbook=xlwt.Workbook(encoding="utf-8")
    for m in range(len(sheet_names)):
        output_sheet=output_workbook.add_sheet(sheet_names[m])
        output_line_counting=0
        for workbook in workbooks:
            sheet=workbook.sheet_by_index(m)
            nrows=sheet.nrows# 行数
            ncols=sheet.ncols# 列数

            for i in range(nrows):
                if(sheet.cell(i,0).value==''):
                    continue
                output_line_counting=write_in_output_sheet(output_sheet,output_line_counting,sheet,i)
    output_workbook.save(output_path+"\\"+"output.xls")
