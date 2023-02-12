import xlrd
book = xlrd.open_workbook('2106006操行分.xlsx')
sheet = book.sheets()[0]
nrows = sheet.nrows
clo20 = sheet.col_values(19)
caoxingfen = []
for i in range(3,nrows):
    caoxingfen.append(clo20[i])
print(caoxingfen)