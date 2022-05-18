import xlrd
import xlwt
import xlutils.copy

data = xlrd.open_workbook('/home/romo/data_operation/data_origin.xls',formatting_info = True)

data_copy = xlutils.copy.copy(data)

sheet_copy = data_copy.get_sheet(0)

sheet = data.sheet_by_index(0)

rows = sheet.nrows

tmp = 0.0

for  i in range(0,rows,1):
    if sheet.cell_type(i,6) == 2 and sheet.cell_type(i,3) != 0:
        tmp = sheet.cell(i,6).value - sheet.cell(i,3).value/86400.0
        sheet_copy.write(i+1,6,tmp)
    elif sheet.cell_type(i,0) ==2 and float(tmp) != 0.0:
        step = tmp + sheet.cell(i,2).value/86400.0
        sheet_copy.write(i,3,step)
        continue

data_copy.save('/home/romo/data_operation/data_origin.xls')

print("check the time now!")

"""
WARNING

1.make sure the type of column(G) is number!
2.after the program is executed,check the time!
3.make sure the path is right!
4.make sure the type of file is ‘xls’！
5.the library need xlrd,xlwt,xlutils!

"""
