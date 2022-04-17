import xlwt

import datetime

workbook = xlwt.Workbook()

worksheet = workbook.add_sheet('My Sheet')

style = xlwt.XFStyle()

style.num_format_str = 'h:mm:ss'

# Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0

worksheet.write(0, 0, datetime.datetime.now(), style)

print(type(datetime.datetime.now()))

workbook.save('Excel_Workbook.xls')

