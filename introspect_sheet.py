from xlrd import open_workbook,cellname,cellnameabs

book = open_workbook('odd.xls')
sheet = book.sheet_by_index(0)

print('sheet_name:%s, sheet_nrows:%s, sheet_ncols:%s'% (sheet.name, sheet.nrows, sheet.ncols))

for row_index in range(sheet.nrows):
    for col_index in range(sheet.ncols):
        print(cellname(row_index, col_index), '-',sheet.cell(row_index,col_index).value)
        print(cellnameabs(row_index, col_index), '-',sheet.cell(row_index,col_index).value)
