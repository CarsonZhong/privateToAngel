#coding=gbk
import xlrd
book = xlrd.open_workbook('data.xlsx')
sheet1 = book.sheets()[0]
nrows = sheet1.nrows
print('���������',nrows)
ncols = sheet1.ncols
print('���������',ncols)
row3_values = sheet1.row_values(2)
print('��3��ֵ',row3_values)
col3_values = sheet1.col_values(2)
print('��3��ֵ',col3_values)
cell_3_3 = sheet1.cell(2,2).value
print('��3�е�3�еĵ�Ԫ���ֵ��',cell_3_3)
