from mmap import mmap,ACCESS_READ
from xlrd import open_workbook

'''Workbooks can be loaded either from a file, an mmap.mmap object or from a string'''

print(open_workbook('simple.xls'))

with open('simple.xls','rb') as f:
    print open_workbook(
        file_contents=mmap(f.fileno(), 0, access=ACCESS_READ)
    )

aString = open('simple.xls','rb').read()
print(open_workbook(file_contents=aString))