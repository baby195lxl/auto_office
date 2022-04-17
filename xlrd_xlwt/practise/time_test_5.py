import xlrd
from xlrd import xldate_as_tuple
import xlwt
import datetime
import math

data = xlrd.open_workbook("time_test.xls")

r_s0 = data.sheet_by_index(0)

wt_b = xlwt.Workbook()

w_s0 = wt_b.add_sheet("time")

time_style = xlwt.XFStyle()

time_style.num_format_str = "h:mm:ss"

for i in range(r_s0.nrows):
    
    if isinstance(r_s0.cell_value(i,6),float) == True:
        
        tup = xldate_as_tuple(r_s0.cell_value(i,6),1)
        print(tup)
        end_time = datetime.time(tup[3],tup[4],tup[5])
        end_time = end_time.strftime("%H:%M:%S.%f")
        end_time = datetime.datetime.strptime(end_time,r"%H:%M:%S.%f")
        print(end_time)
    
        if r_s0.cell_value(i,3) >= 60.0:
           
            mint = int(r_s0.cell_value(i,3) // 60)
            sec = r_s0.cell_value(i,3) - 60.0 * mint
            tmp = math.modf(sec)
            sec = int(tmp[1])
            microsec = round(tmp[0],6)
            sped_time = datetime.time(0,mint,sec,int(microsec * 1000000))
            sped_time = sped_time.strftime("%H:%M:%S.%f")
            sped_time = datetime.datetime.strptime(sped_time,"%H:%M:%S.%f")
            print(sped_time)
    
        else:            
            
            tmp = math.modf(r_s0.cell_value(i,3))
            sec = int(tmp[1])
            microsec = round(tmp[0],6)
            sped_time = datetime.time(0,0,sec,int(microsec * 1000000))
            sped_time = sped_time.strftime("%H:%M:%S.%f")
            sped_time = datetime.datetime.strptime(sped_time,"%H:%M:%S.%f")
            print(sped_time)

        star_time = end_time - sped_time
        print(star_time)

#        star_time = str(star_time)
#        print(star_time)
#        w_s0.write(i+1,6,star_time,time_style) 
    
    else:
        continue

for j in range(r_s0.nrows):
  
    if isinstance(r_s0.cell(j,0),float) == True:
  
        if r_s0.cell_value(j,3) >= 60.0:
  
            mint = int(r_s0.cell_value(j,3) // 60)
            sec = r_s0.cell_value(j,3) - 60.0 * mint
            tmp = math.modf(sec)
            sec = int(tmp[1])
            microsec = round(tmp[0],6)
            step_time = datetime.time(0,mint,sec,int(microsec * 1000000))
            step_time = step_time.strftime("%H:%M:%S.%f")
            step_time = datetime.datetime.strptime(step_time,"%H:%M:%S.%f")
            print(step_time)

        else:

            tmp = math.modf(r_s0.cell_value(j,3))
            sec = int(tmp[1])
            microsec = round(tmp[0],6)
            step_time = datetime.time(0,0,sec,int(microsec * 1000000))
            step_time = step_time.strftime("%H:%M:%S.%f")
            step_time = datetime.datetime.strptime(step_time,"%H:%M:%S.%f")
            print(step_time)



#wt_b.save("time_result.xls")
