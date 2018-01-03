from xlrd import open_workbook,error_text_from_code
'''error_text_from_code is a dictionary. Can be used to turn error codes into error messages
{0: '#NULL!', 36: '#NUM!', 7: '#DIV/0!', 42: '#N/A', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?'}
'''
book = open_workbook('types.xls')
sheet = book.sheet_by_index(0)

print error_text_from_code
print sheet.cell(5,2).value
print sheet.cell(5,3).value
print error_text_from_code[sheet.cell(5,2).value]
print error_text_from_code[sheet.cell(5,3).value]