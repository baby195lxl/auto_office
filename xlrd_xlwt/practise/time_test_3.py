import xlrd
from xlrd import xldate_as_tuple
import xlwt
import datetime


data = xlrd.open_workbook("time_test.xls")

r_s0 = data.sheet_by_index(0)

tup_1 = xldate_as_tuple(r_s0.cell_value(2,6),1)

print(tup_1)

init_time =datetime.time(tup_1[3],tup_1[4],tup_1[5])

print(type(init_time))

init_time = init_time.strftime("%H:%M:%S.%f")
init_time = datetime.datetime.strptime(init_time,r"%H:%M:%S.%f")

print(init_time)
print(type(init_time))

sped_time = datetime.time(0,0,9,508789)

print(sped_time)
print(type(sped_time))

sped_time = sped_time.strftime("%H:%M:%S.%f")
sped_time = datetime.datetime.strptime(sped_time,r"%H:%M:%S.%f")

print(sped_time)
print(type(sped_time))

star_time = init_time - sped_time

print(star_time)
print(type(star_time))

star_time = str(star_time)

#star_time = star_time.strftime("%H:%M:%S.%f")
star_time = datetime.datetime.strptime(star_time,r"%H:%M:%S.%f")

print(star_time)
print(type(star_time))


wb = xlwt.Workbook()

style = xlwt.XFStyle()

style.num_format_str = "h:mm:ss"

w_s0 = wb.add_sheet("mts")

w_s0.write(3,6,star_time,style)

wb.save("time_test_save.xls")



