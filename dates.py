from datetime import date, datetime, time
from xlrd import open_workbook, xldate_as_tuple

'''xldate_as_tuple method is provided for turning the float in Date cell into a tuple suitable 
for instantiating various date/time objects.'''

book = open_workbook('types.xls')
sheet = book.sheet_by_index(0)

# book.datemode object 'book's type
date_value=xldate_as_tuple(sheet.cell(3,2).value, book.datemode)
print datetime(*date_value), date(*date_value[:3])

datetime_value=xldate_as_tuple(sheet.cell(3,3).value, book.datemode)
print datetime(*datetime_value)

time_value=xldate_as_tuple(sheet.cell(3,4).value, book.datemode)
print time(*time_value[3:])
print datetime(*time_value)