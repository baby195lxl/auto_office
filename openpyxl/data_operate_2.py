import xlrd,xlwt

data = xlrd.open_workbook('/home/romo/data_operation/data_origin.xls',formatting_info = True)

sheet = data.sheet_by_index(0)

wb = xlwt.Workbook()

s0 = wb.add_sheet('MTS')

headers = ['力(KN)','位移(mm)','步长(s)','时间']
for x in headers:
    s0.write(0,headers.index(x),x)

rows = sheet.nrows

cols = sheet.ncols

for i in range(0,cols,1):
    for j in range(0,rows,1):
        if sheet.cell_type(j,0) == 2:
            s0.write(j,i,sheet.cell(j,i).value)
        else:
            continue

wb.save('data_finished.xls')
