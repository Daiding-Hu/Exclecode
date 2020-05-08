import openpyxl
import xlrd
import xlwt
import datetime
import pymssql

workbook = xlrd.open_workbook('练习.xlsx')
sheet_names = workbook.sheet_names()
print(sheet_names)
sheets = workbook.sheets()  # 获取sheet对象，赋给sheets
print(sheets)
sheet1 = workbook.sheet_by_index(0)  # 获取sheet中第一张表，0代表第一张
print(sheet1)
sheet2 = workbook.sheet_by_name("表格2")  # 获取表格文件中名字为‘表格2’的sheet对象给sheetn
print(sheet2)
sheet_is_load = workbook.sheet_loaded('表格2')  # 通过表格名称判断sheet是否导入
print(sheet_is_load)
sheet_is_load2 = workbook.sheet_loaded(0)  # 通过index判断sheeet是否导入
print(sheet_is_load2)

'''对行的操作'''
nrows = sheet1.nrows  # 获取sheet1中的有效行数
print(nrows)
row_value = sheet1.row_values(rowx=2)  # 获取第2行的数据存放到列表中
print(row_value)
row_value1 = sheet1.row_values(rowx=2, start_colx=1, end_colx=5)  # 截取第二行数据的部分数据，从0开始，下标左可以取到，右不可以
print(row_value1)
row_object = sheet1.row(rowx=2)  # 获取第2行的单元对象，没一个单元格数据的类型，
# 单元类型：empty，string，number，date，boolean， error
print(row_object)
row_type = sheet1.row_types(rowx=2)  # 获取第二行的单元类型
# 单元类型ctype：empty为0，string为1，number为2，date为3，boolean为4， error为5；
print(row_type)
row_len = sheet1.row_len(rowx=2)
print(row_len)
'''对列的操作'''
ncols = sheet1.ncols  # sheet1中有多少有效列
print(ncols)
ncols_value = sheet1.col_values(colx=1)
print(ncols_value)  # 第一列的值
ncols_value1 = sheet1.col_values(4, 1, 3)  # 第四列中1，3行的值
print(ncols_value1)
cols_slic = sheet1.col_slice(colx=2)  # 第二列的数据和类型
print(cols_slic)
cols_type = sheet1.col_types(colx=2)  # 第二列的单元格类型对应编号
print(cols_type)
'''对sheet对象中的单元执行操作'''
cell_value = sheet1.cell(rowx=0, colx=0)  # 0行0列的数值和类型
print(cell_value)
a = sheet1.cell_value(0, 0)  # 0行0列的数值
print(a)
a1 = sheet1.cell_type(1, 13)
print(1)
b1 = sheet1.cell_value(1, 13)
print(b1)
c1 = xlrd.xldate.xldate_as_datetime(b1, workbook.datemode)  # 提取时间
print(c1)
d1 = c1.strftime('%Y/%m/%d')  # 转换时间格式
print(d1)
'''合并单元格'''
''' 获取合并的单元格
若表格为xls格式的，打开workbook时需将formatting_info设置为True，然后再获取sheet中的合并单元格；
若表格有xlsx格式的，打开workbook时保持formatting_info为默认值False，然后再获取sheet中的合并单元格；
workbook1 = xlrd.open_workbook("测试.xls", formatting_info=True)'''
a2 = sheet2.merged_cells  # 获取xlsx格式的excel文件中的合并单元格的位置
print(a2)
b2 = sheet2.cell_value(2, 8)  # 读取合并单元格数据（仅需“起始行起始列”即可获取数据）
print(b2)
for (row_start, row_end, col_start, col_end) in sheet2.merged_cells:  # 使用for循环获取所有的合并单元格数据
    print(sheet2.cell_value(row_start, col_start))
