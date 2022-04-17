import math
import xlrd

data = xlrd.open_workbook("time_test.xls")

r_s0 = data.sheet_by_index(0)

sec = r_s0.cell_value(2,3)

if sec >= 60.0:
    mint = int(sec // 60)
    sec = sec - 60 * mint
    tmp =math.modf(sec)
    sec = int(tmp[1])
    microsec = round(tmp[0],6)
    print(mint)
    print(sec)
    print(microsec)
else:
    tmp = math.modf(sec)
    sec = int(tmp[1])
    microsec = round(tmp[0],6)
    print(sec)
    print(microsec)
