import xlrd
book = xlrd.open_workbook('截止第18周末2106006班公益分表.xlsx')
sheet = book.sheet_by_index(0)
clo57 = sheet.col_values(sheet.merged_cells[1][3])
nrows = sheet.nrows
gongyifen = []
tep = []
merged=sheet.merged_cells
for i in range(0,len(merged)):
    if(merged[i][3]==57 and merged[i][2]==56):
       tep.append(merged[i]) 
    
for i in range(0,len(tep)):
    gongyifen.append(sheet.cell_value(tep[i][0],57))
print(gongyifen)