# coding=gbk
import sys

# import settings
srcfile='f:/template.xls'
dstfile='f:/file.xls'

import os,shutil
def mycopyfile(srcfile,dstfile):
    if not os.path.isfile(srcfile):
        print ("%s not exist!"%(srcfile))
        return False
    else:
        fpath,fname=os.path.split(dstfile)    #分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)                #创建路径
        shutil.copyfile(srcfile,dstfile)      #复制文件
        print ("copy %s -> %s"%( srcfile,dstfile))
    
    return True

FILENAME = 'f:/出货明细.txt'

try:
    fi = open(FILENAME)
except:
    print("{}{}{}".format("file name with ", FILENAME, " not exist"))
    exit()


lines = fi.readlines()
# data = []
context = []

# context = [
#     { # item
#         订购日期 : 2018/07/19 11:29
#         产品型号 : 100PACK-MSTH6-20
#     },
#     { # item
#         订购日期:2018/07/19 11:29
#         产品型号:UTSNA1-30
#     }
# ]
# new_start = True
item = {}
for line in lines:
#     print(line)
    
    if "[订购日期]" in line:
#         print(line)
        if item:
            context.append(item) # 将前一个单子保存
        item = {}
        item['date'] = line # 初始化
    if "[产品型号]" in line:
        try:
            line = line.split(":")[1]
        except:
            print("{}".format(line, " 格式不准确"))
        item["model"] = line
    if "[数量]" in line:
        try:
            line = line.split(":")[1]
        except:
            print("{}".format(line, " 格式不准确"))
        item["number"] = line
if item:
    context.append(item)
print(context)

mycopyfile(srcfile,dstfile)

sheetname = "收料单格式"

STARTING_LINE_ROW_OFFSET = 9
INDEX_OFFSET = 0
PICTURE_OFFSET = INDEX_OFFSET + 1
NAME_OFFSET = PICTURE_OFFSET + 1
NUMBER_OFFSET = NAME_OFFSET + 1

# from openpyxl import load_workbook
# wb = load_workbook(dstfile)
# ws = wb[sheetname]
# 
# ws.cell(0, 0, "test")
# wb.save()

# #两个for循环遍历整个excel的单元格内容
# for i,row in enumerate(ws.iter_rows()):
#     for j,cell in enumerate(row):
#         ws2.cell(row=i+1, column=j+1, value=cell.value)

# https://blog.csdn.net/u013176681/article/details/51119071      
import xlwt, xlrd
from xlutils.copy import copy
 
from datetime import datetime

# https://www.cnblogs.com/xiaodingdong/p/8012282.html
style = xlwt.XFStyle()
borders = xlwt.Borders()
borders.bottom = xlwt.Borders.THIN # 下面加横线
borders.left = xlwt.Borders.THIN
font = xlwt.Font()
font.name = "Times New Roman"
font.bold = "on"
# font.color-index = "red"

style.borders  = borders 
style.font  = font

 
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
 
rb = xlrd.open_workbook(dstfile,formatting_info=True)
wb = copy(rb)  
ws = wb.get_sheet(1)
for i,line in enumerate(context):
    ws.write(STARTING_LINE_ROW_OFFSET + i, NAME_OFFSET, line.get('model', None), style)
    ws.write(STARTING_LINE_ROW_OFFSET + i, NUMBER_OFFSET, line.get('number', None), style)
#     ws.write(STARTING_LINE_ROW_OFFSET + i, NAME_OFFSET, line.get('model', None))
    
#     ws.write(1, 0, datetime.now(), style1)
#     ws.write(2, 0, 1)
#     ws.write(2, 1, 1)
#     ws.write(2, 2, xlwt.Formula("A3+B3"))
#     ws.write(1, 6, 'changed!') 
wb.save(dstfile)


fi.close()