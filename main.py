import xlrd
import sys
from datetime import datetime
import time
from xlrd import open_workbook,xldate_as_tuple
 
filename = 'tel.xls'

vcf_template = """
BEGIN:VCARD
VERSION:3.0
N:;{name};;;
FN:{name} 
TEL;type=CELL:{mobile}
TEL;type=HOME:
TEL;type=WORK:
EMAIL;type=INTERNET;type=WORK;type=pref:
EMAIL;type=INTERNET;type=HOME;type=pref:
EMAIL;type=INTERNET;type=CELL;type=pref:
item1.ADR;type=WORK:;; 
item2.ADR;type=HOME;type=pref:;; 
item3.URL;type=pref:
END:VCARD
"""
 
try:
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_name('tel')
 
    cell = sheet.cell(1,1)
    print(cell)
    print(cell.value)
    print(cell.ctype)
    if cell.ctype == xlrd.XL_CELL_DATE:
        date_value = xldate_as_tuple(cell.value,book.datemode)
        print(datetime(*date_value))
 
    dataset = []
    for r in range(sheet.nrows):
        col = []
        for c in range(sheet.ncols):
            col.append(str(sheet.cell(r,c).value).split('.')[0])
        dataset.append(col)
    i = 0
    with open("output.vcf", "w") as f:
        for person in dataset:
            if(len(person[1]) < 7):
                continue
            vcf = vcf_template.format(name=person[0], mobile=person[1])
            f.write(vcf)
            i = i + 1
            print(vcf)
    print("count = %d" % i)
 
except xlrd.XLRDError as e:
    print('Read excel error:%s' % e)
    sys.exit(1)