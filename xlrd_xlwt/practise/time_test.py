import xlrd
from xlrd import xldate_as_tuple
import xlwt
import datetime

data = xlrd.open_workbook("time_test.xls")

#print(data.sheet_loaded(0))

#print(data.sheet_names())

r_s0 = data.sheet_by_index(0)

#print("rows = %d" % r_s0.nrows)

#print("cols = %d" % r_s0.ncols)

#print(r_s0.cell(2,6).value)

#print("type = %d" % r_s0.cell_type(2,6))

"""
wb = xlwt.Workbook()

w_s0 = wb.add_sheet("MTS_TIME")

for i in range(r_s0.nrows):
    val = r_s0.cell_value(i,6)
    w_s0.write(i,0,val)

wb.save("time_test_save.xls")
"""
#print(r_s0.row_types(2))

print("type of cell(2,6) = %s" % type(r_s0.cell(2,6).value))

tmp_1 = xldate_as_tuple(r_s0.cell(2,6).value,0)

print("type of temp1 = %s" % type(tmp_1))

wb = xlwt.Workbook()

w_s0 = wb.add_sheet("MTS")




