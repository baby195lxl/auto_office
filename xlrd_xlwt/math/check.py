import xlrd
import xlwt

data = xlrd.open_workbook("time_test.xls")

wt_b = xlwt.Workbook()

r_s0 = data.sheet_by_index(0)

w_s0 = wt_b.add_sheet("checK sheet")

j = 0

for i in range(r_s0.nrows):

    if isinstance(r_s0.cell_value(i,0),float) == True:
        
        w_s0.write(j,0,r_s0.cell_value(i,0))
                
        j += 1
    
    else:
        continue

wt_b.save("check.xls")
