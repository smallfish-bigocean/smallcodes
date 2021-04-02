import requests
from bs4 import BeautifulSoup
from  openpyxl import  Workbook 
from openpyxl import load_workbook
import winreg

#额外任务，获取当前用户的桌面路径，这样就可以把抓到的Excel表格存到桌面了
key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
mydesktop = winreg.QueryValueEx(key,"Desktop")[0]
mydesktop = mydesktop.replace('\\', '\\\\')
#获取MTM编号
mtm = input('请输入MTM Number: ')
url = 'https://psref.lenovo.com/Search?kw='+mtm
#验证MTM是否存在
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko)'
'Chrome/51.0.2704.63 Safari/537.36'}
res = requests.get(url,headers=headers)
html = res.text
soup = BeautifulSoup( html,'html.parser')
item = soup.find('p',class_='filtered_title')
try:
    if item.text != '':
        print('找到了此MTM')
    #如果找到了MTM，开始抓取MTM配置
    #先得找MTM页面地址
    item = soup.find('ul',class_='modets_list')
    #拼接MTM地址
    mtmurl = 'https://psref.lenovo.com/'+item.attrs['modeldetaillinkpart'].replace('{ModelCode}',mtm)
    #打开MTM地址，获取配置信息
    res = requests.get(mtmurl,headers=headers)
    html = res.text
    soup = BeautifulSoup( html,'html.parser')
    table = soup.find('table',class_='SpecValueTable')
    #开始写入Excel表格
    wb = Workbook()
    ws = wb.active
    lists = []
    for tr in table.find_all('tr'):
        for td in tr.find_all('td'):
            lists.append(td.text)
        ws.append(lists)
        lists.clear()
    wb.save(mydesktop+'\\\\MTM_Config.xlsx')

except AttributeError:
    print('未发现此MTM， 请检查输入的MTM是否正确')
