import requests
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

def check(rul):
    try:
        code=requests.get(rul).status_code
    except Exception as err:
        err='异常'
        return  err
    return  code

file_name='weburl.xlsx'
wb=openpyxl.load_workbook(file_name)
sheet=wb['web']

for row in range(2,sheet.max_row+1):
    checkurl=sheet['C'+str(row)].value
    sheet['D'+str(row)].value=check( checkurl )

wb.save(file_name)



recode=check('http://www.jb51.net')
print(recode)
recode=check('http://www.glswxgj.gov.cn/')
print(recode)
