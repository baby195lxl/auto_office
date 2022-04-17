import xlrd
from xlrd import xldate_as_tuple

data = xlrd.open_workbook("time_test.xls")

r_s0 = data.sheet_by_index(0)

print(r_s0.cell_value(6,6))

print(type(r_s0.cell_value(6,6)))

print(isinstance(r_s0.cell_value(6,6),float))
