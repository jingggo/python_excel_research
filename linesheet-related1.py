from xlrd import open_workbook
from xlwt import Workbook

def getTemplateUploadLinesheet():
    book = open_workbook(r'E:\RFScripts\linesheet_upload_data_ef\Linesheet upload use case-ef.xlsx')
    sheet = book.sheet_by_name('EF-DATA')
    return sheet.row_values(3)

def getExistingLineSheet(nrows,xrow,xcol=1):
    book = open_workbook(r'E:\RFScripts\linesheet_upload_data_ck\Linesheet upload use case.xlsx')

    sheet0 = book.sheet_by_index(0)
    rlist = []
    for row in range(nrows):
        rlist.append(sheet0.row_values(xrow+row, xcol, 9))
    # sheet1 = book.sheet_by_index(1)

    # print sheet0.row_slice(4, 1, 7)
    # print sheet0.row_values(4, 1, 7)
    # print sheet0.row_types(4, 1, 7)
    return (rlist)

def getUploadLineSheet(nrows,xrow,xcol=1):
    book = open_workbook(r'E:\RFScripts\linesheet_upload_data_ck\Linesheet upload use case.xlsx')

    sheet0 = book.sheet_by_index(0)
    rlist = []
    for row in range(nrows):
        rlist.append(sheet0.row_values(xrow+row, xcol, 9))
    return (rlist)

def createExcel():
    book = Workbook()
    sheet1 = book.add_sheet('EF-DATA', cell_overwrite_ok=True)

    scenario_indexs = [2,11,23,31,43,55,67,75,87,99,114,126,135,145,154,169,184,200,215,230,246,254,266,278,286,298,310,322,330,338,350,359,368,380,388,396,408,417,426,434,444,452,462,472,480,492,500,512,520,528,538,546,556]
    existing_counts =[2,1, 1, 3, 4, 1, 1, 3, 4, 4, 4, 2, 2, 1, 6, 6, 6, 6, 6, 6, 1, 3, 4, 1, 3, 4, 3, 1, 1, 3, 2, 2, 3, 1, 1, 3, 2, 2, 1, 2, 1, 2, 2, 1, 3, 1, 3, 1, 1, 2, 1, 2, 2]
    upload_counts =[1,5, 1, 3, 2, 5, 1, 3, 2, 5, 2, 1, 2, 2, 3, 3, 4, 3, 3, 4, 1, 3, 2, 1, 3, 2, 3, 1, 1, 3, 1, 1, 3, 1, 1, 3, 1, 1, 1, 2, 1, 2, 2, 1, 3, 1, 3, 1, 1, 2, 1, 2, 2]
    scenario_names = getScenario(0)
    row_index = 0
    for index,existing_count,upload_count,scenario_name in zip(scenario_indexs,existing_counts,upload_counts,scenario_names):
        existing_values, upload_values = getExpectedLinesheets(index+2,existing_count,upload_count)#(4,2,1)
        # print(existing_values, upload_values)
        row_index = writeLineToSheet(sheet1,row_index,existing_values,upload_values,scenario_name)
        row_index += 2
        # print(row_index)
        # break

    book.save(r'E:\RFScripts\linesheet_upload_data_ef\test4.xls')

def writeLineToSheet(sheet1, row_index, existing_values,upload_values,scenario_name):
    xrow = row_index
    sheet1.write(xrow, 0, scenario_name)
    xrow += 1

    sheet1.write(xrow, 0, 'Existing Record :')
    xrow += 1

    template_linesheet_title = getTemplateUploadLinesheet()
    [sheet1.write(xrow, ic, value) for ic, value in enumerate(template_linesheet_title)]
    xrow += 1

    for irow in range(len(existing_values)):
        row_values = existing_values[irow]
        for xcol, value in enumerate(row_values):
            sheet1.write(irow + xrow, xcol, value)
    xrow += len(existing_values)

    sheet1.write(xrow, 0, 'Upload File :')
    xrow += 1

    [sheet1.write(xrow, ic, value) for ic, value in enumerate(template_linesheet_title)]
    xrow += 1

    '''upload linesheet'''
    for iurow in range(len(upload_values)):
        row_upload_values = upload_values[iurow]
        for xcol, value in enumerate(row_upload_values):
            sheet1.write(iurow + xrow, xcol, value)
    xrow += len(upload_values)

    sheet1.write(xrow, 0, 'Expected Result:')
    xrow += 1

    sheet1.write(xrow, 0, 'Actual Result:')
    return xrow

def writeLineSheet(existing_values, upload_values):
    book = Workbook()
    sheet1 = book.add_sheet('EF-DATA', cell_overwrite_ok=True)
    xrow=0
    sheet1.write(xrow, 0, 'Scenario 1 : New Style')
    xrow += 1

    sheet1.write(xrow, 0, 'Existing Record :')
    xrow += 1

    template_linesheet_title = getTemplateUploadLinesheet()
    [sheet1.write(xrow, ic, value) for ic, value in enumerate(template_linesheet_title)]
    xrow += 1

    for irow in range(len(existing_values)):
        row_values = existing_values[irow]
        for xcol, value in enumerate(row_values):
            sheet1.write(irow+xrow,xcol,value)
    xrow += len(existing_values)

    sheet1.write(xrow, 0, 'Upload File :')
    xrow += 1

    [sheet1.write(xrow, ic, value) for ic, value in enumerate(template_linesheet_title)]
    xrow += 1

    '''upload linesheet'''
    for iurow in range(len(upload_values)):
        row_values = existing_values[iurow]
        for xcol, value in enumerate(row_values):
            sheet1.write(iurow+xrow, xcol, value)
    xrow += len(upload_values)

    sheet1.write(xrow, 0, 'Expected Result:')
    xrow += 1

    sheet1.write(xrow, 0, 'Actual Result:')
    # sheet1.write(0,0,'A1')
    # sheet1.write(0,1,'B1')
    # row1 = sheet1.row(1)
    # row1.write(0,'A2')
    # row1.write(1,'B2')
    book.save(r'E:\RFScripts\linesheet_upload_data_ef\test.xls')
    return xrow

def addToRealLinesheet(linesheet):
    tempLinesheet=['','','','','','AUTOTEST','','','','','','143.0','288.0','','','','','EILEENFISHER','WOMENSWEAR','MISSY','AUTOTESTSTORY','','AAH - SLEEK TENCEL','','W - SWEATERS','','','','SWEATERS','M7 (XXS-XXL)','FALL','2018','0.0','0.0','','','','','','DOMESTIC','67.76','','','','','','','','','','','','','','M','','','','','','','','','','','','','F','','0.0','0.0','','','','','','','0','','0.0','0','','M4','ANDARI FASHION, INC.','UNITED STATES','','']
    '''Corporate,style,upc,color code,color des, size code,size desc'''
    if not linesheet[0] == '':
        tempLinesheet.pop(17)
        tempLinesheet.insert(17,'EILEENFISHER')
    if not linesheet[1] == '':
        tempLinesheet.pop(0)
        tempLinesheet.insert(0, 'EF'+linesheet[1])
    if not linesheet[2] == '':
        tempLinesheet.pop(1)
        tempLinesheet.insert(1, linesheet[2])
    if not linesheet[3] == '':
        tempLinesheet.pop(6)
        tempLinesheet.insert(6, linesheet[3])
    if not linesheet[4] == '':
        tempLinesheet.pop(7)
        tempLinesheet.insert(7, linesheet[4])
    if not linesheet[5] == '':
        tempLinesheet.pop(8)
        tempLinesheet.insert(8, linesheet[5])
    if not linesheet[6] == '':
        tempLinesheet.pop(9)
        tempLinesheet.insert(9, linesheet[6])
    if not linesheet[7] == '':
        if linesheet[7]=='A/M':
            linesheet[7]='M'
        tempLinesheet.pop(54)
        tempLinesheet.insert(54, linesheet[7])
    return tempLinesheet

def getExpectedLinesheets(existing_row_index, existing_nrows, upload_nrows):
    src_existing_linesheets = getExistingLineSheet(existing_nrows,existing_row_index)#2 rows for existing, the row_index is 4
    src_upload_linesheets = getUploadLineSheet(upload_nrows,existing_row_index+existing_nrows+2)
    # print(src_existing_linesheets)
    # print(src_upload_linesheets)

    existing_real_linesheet = []
    for src_existing_linesheet in src_existing_linesheets:
        existing_real_linesheet.append(addToRealLinesheet(src_existing_linesheet))

    upload_real_linesheet = []
    for src_upload_linesheet in src_upload_linesheets:
        upload_real_linesheet.append(addToRealLinesheet(src_upload_linesheet))
    return existing_real_linesheet, upload_real_linesheet

def getScenario(sheet_index):
    scenario_indexs = [1, 10, 22, 30, 42, 54, 66, 74, 86, 98, 113, 125, 134, 144, 153, 168, 183, 199, 214, 229, 245, 253, 265, 277, 285, 297, 309, 321, 329, 337, 349, 358, 367, 379, 387, 395, 407, 416, 425, 433, 443, 451, 461, 471, 479, 491, 499, 511, 519, 527, 537, 545, 555]
    book = open_workbook(r'E:\RFScripts\linesheet_upload_data_ck\Linesheet upload use case.xlsx')
    sheet = book.sheet_by_index(sheet_index)
    scenario_names=[]
    for scenario_index in scenario_indexs:
        scenario_names.append(sheet.cell(scenario_index,1).value)
    # print scenario_names
    return scenario_names
createExcel()