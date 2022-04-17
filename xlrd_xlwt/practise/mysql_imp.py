import xlrd
import xlwt

data = xlrd.open_workbook('test_excel.xls')

r_s0 = data.sheet_by_index(0)

#array_0 = r_s0.col_values(0)
#array_1

#print(array_1)

wb = xlwt.Workbook()

w_s0 = wb.add_sheet('MTS')

for j in range(0,4):
    for i in range(r_s0.nrows):
        val = r_s0.cell_value(i,j)
        w_s0.write(i,j,val)

wb.save("test_excel_save.xls")
