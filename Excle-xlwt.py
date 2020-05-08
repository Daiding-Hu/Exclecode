import xlwt

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('表格1')
# worksheet.write(0, 0, 'this is test')  # 在0，0处写入数据
# workbook.save('练习2.xlsx')
list_num = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
for i in range(len(list_num)):
    worksheet.write(0, i, list_num[i])
style = xlwt.XFStyle()  # 初始化样式
font = xlwt.Font()  # 创建字体
font.name = 'Times New Roman'
font.bold = True
workbook.save('练习2.xlsx')
