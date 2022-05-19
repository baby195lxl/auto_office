import xlrd

data = xlrd.open_workbook('/home/romo/test/test_xlread.xls') #open the workbook

#print(data.sheet_loaded(0)) # load '0' sheet

#data.unload_sheet(0) # unload '0' shhet

#print(data.sheet_loaded(0))

#print(data.sheets()) # show all sheet

#print(data.sheets()[0]) # show '0' sheet

#print(data.sheet_by_index(0)) # show sheet by the number of sheet

#print(data.sheet_by_name('first')) # show sheet by the name of sheet

#print(data.sheet_names()) # show all name of sheets

#print(data.nsheets) # show how many sheet in the workbook

sheet = data.sheet_by_index(0) # get the sheet(0)

########################################## operate the rows ###########################################

#print(sheet.nrows) #show how many rows does the sheet(0) have

#rows = int(sheet.nrows)

#print("num of rows = %d" % rows)

#print(sheet.row(0)) #show the content of row(0)

#print(sheet.row_types(1)) #show the contents' types of row(1) [0:empty 1:str 2:num 3:date 4:bool 5:error]

#print(sheet.row(1)[2]) #show the content of cell(1)[2]

#print(sheet.row(1)[2].value) #only show the value of cell(1)[2]

#print(sheet.row_values(1)) #only show the value of row(1)

#print(sheet.row_len(1)) #show the length of row(1)

######################################## operate the columns ########################################

#print(sheet.ncols) #show how many cols does the sheet(0) has

#print(sheet.col_values(1)) #show all value of col(1)

#print(sheet.col(1)[2]) #show the content of cell(2)[1]

#print(sheet.col_types(1)) #show the contents' types of col(1)

###################################### operate the cells ###########################################

#print(sheet.cell(1,2)) #show the content of cell(1)[2]

#print(sheet.cell_type(1,2)) #show the type of cell(1)[2]

#print(sheet.cell(1,2).ctype) #same as ".cell_type('','')"

#print(sheet.cell(1,2).value) #show the value of cell(1)[2]

#print(sheet.cell_value(1,2)) #same os "cell('','').value"




