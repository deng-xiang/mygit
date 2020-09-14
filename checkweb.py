import requests
import openpyxl

from bs4 import BeautifulSoup
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
#    sheet['D'+str(row)].value=check( checkurl )
    print(str(sheet['B'+str(row)].value) +':'+ str(sheet['D'+str(row)].value) )

wb.save(file_name)





url='http://seo.chinaz.com/'
postdata = {
"q":"www.guilin.gov.cn"
}
r = requests.post(url, postdata)
#print (r.text)

soup = BeautifulSoup(r.text, 'html.parser')
#网站域名

url=soup.find('div',attrs={'class':'ball color-63'})
print("网站域名:"+url.string)
#备案信息
document=soup.find('td',attrs={'class':'_chinaz-seo-newtc _chinaz-seo-newh40'})
#print(document.span.a)
document=document.find('span')
document=document.find('i')
document=document.find('a')
print("备案号:"+document.string)
#IP
IP=soup.find('td',attrs={'class':'_chinaz-seo-newh78 _chinaz-seo-newinfo'})

IP=IP.find('div',attrs={'class':'pb5'})
IP=IP.find('span')
IP=IP.find('i')
IP=IP.find('a')
ip1=IP.string.split('[')[0]
addr=IP.string.split('[')[1].replace(']','')
print("IP:"+ip1)
print("物理位置:"+addr)








