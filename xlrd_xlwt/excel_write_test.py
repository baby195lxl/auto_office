import xlwt

####################### set title style ###############################################

title_style_type = xlwt.XFStyle() #initial the style of title

######################## set title font ############################################### 

title_font = xlwt.Font()
title_font.name = "Times New Roman" #set the style of words is "Times New Roman"
title_font.bold = False #set the words bold is "unbolded"
title_font.height = 11*20 #set the size of words is "11"
title_font.colour_index = 0x08 #set the colour of words is "black"
title_style_type.font = title_font

###################### set titles' cell alignment ############################################

cell_align = xlwt.Alignment()
cell_align.horz = xlwt.Alignment.HORZ_CENTER
cell_align.vert = xlwt.Alignment.VERT_CENTER
title_style_type.alignment = cell_align

##################### set the border of title ################################################

cell_border = xlwt.Borders()
cell_border.left = xlwt.Borders.NO_LINE
cell_border.right = xlwt.Borders.NO_LINE
cell_border.top = xlwt.Borders.NO_LINE
cell_border.bottom = xlwt.Borders.NO_LINE
title_style_type.borders = cell_border
#############################################################################################
#############################################################################################

################### set serial_num style #######################################################

serial_num_style = xlwt.XFStyle()

################## set serial_num background ##################################################

bg_colour = xlwt.Pattern()
bg_colour.pattern = xlwt.Pattern.SOLID_PATTERN
bg_colour.pattern_fore_colour = 22  
serial_num_style.pattern = bg_colour

################################################################################################
################################################################################################

wb = xlwt.Workbook() #create the workbook

s0 = wb.add_sheet("first") #create a sheet which named first

s0.write_merge(11,11,0,9,"total",title_style_type) #merge the cells that we need [(r1,r2,c1,c2,lable="",style=)

data = (("序号","类型","规格(mm)","长(m)","宽(m)","截面积(m2)","长度(m)","体积(m3)","数量(个)","总体积(m3)","重量(吨)"),(1,"热轧H型钢","HW125*125*6.5*9",'','',0.003031,1.735,0.005258785,2,0.01051757,0.082562925)) #def the data that we need to input

for i,item in enumerate(data):
    for j,val in enumerate(item):
        if j == 0:
            s0.write(i,j,val,serial_num_style)
        else:
            s0.write(i,j,val)              # input the data

s1 = wb.add_sheet("image") #create a sheet which name image

s1.insert_bitmap("test_image.bmp",0,0) #insert a image into s1 [(filename,row.col,x = 0,y = 0,scale_x = 1,scale_y = 1)]



wb.save("excel_wt_test.xls") #save the workbook














