import xlrd
from xlrd import xldate_as_tuple
import xlwt
import datetime

data = xlrd.open_workbook("time_test.xls")

r_s0 = data.sheet_by_index(0)

#tmp_1 = xldate_as_tuple(r_s0.cell(2,6).value,0)

#print(tmp_1)

wb = xlwt.Workbook()

style = xlwt.XFStyle()

style.num_format_str = "h:mm:ss"

w_s0 = wb.add_sheet("MTS")

#val = xldate_as_tuple(float(r_s0.cell(2,6).value),0)

tmp_1 = xldate_as_tuple(r_s0.cell(2,6).value,0)

print(tmp_1)

tmp_2 = list(tmp_1)

print(type(tmp_2))

print(tmp_2)

#w_s0.write(0,0,tmp_1,style)

#wb.save("time_test_save.xls")
