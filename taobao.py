import os
import re
import xlwt
import sqlite3
import requests
from win32.win32crypt import CryptUnprotectData

def getcookiefromchrome():
    host = '.taobao.com'
    cookies_str = ''
    cookiepath=os.environ['LOCALAPPDATA']+r"\Google\Chrome\User Data\Default\Cookies"
    sql="select host_key,name,encrypted_value from cookies where host_key='%s'" % host
    with sqlite3.connect(cookiepath) as conn:
        cu=conn.cursor()        
        cookies={name:CryptUnprotectData(encrypted_value)[1].decode() for host_key,name,encrypted_value in cu.execute(sql).fetchall()}
        for key,values in cookies.items():
                cookies_str = cookies_str + str(key)+"="+str(values)+';'
        return cookies_str

def writeExcel(ilt,name):
    if(name != ''):
        count = 0
        workbook = xlwt.Workbook(encoding= 'utf-8')
        worksheet = workbook.add_sheet('temp')
        worksheet.write(count,0,'序号')
        worksheet.write(count,1,'购买')
        worksheet.write(count,2,'价格')
        worksheet.write(count,3,'描述')
        for g in ilt:
            count = count + 1
            worksheet.write(count,0,count)
            worksheet.write(count,1,g[0])
            worksheet.write(count,2,g[1])
            worksheet.write(count,3,g[2])
        workbook.save(name+'.xls')
        print('已保存为：'+name+'.xls')
    else:
        printGoodsList(ilt)

def getHTMLText(url):
    cookies = getcookiefromchrome()
    kv = {'cookie':cookies,'user-agent':'Mozilla/5.0'}
    try:
        r = requests.get(url,headers=kv, timeout=30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return ""
     
def parsePage(ilt, html):
    try:
        plt = re.findall(r'\"view_price\"\:\"[\d\.]*\"',html)
        tlt = re.findall(r'\"raw_title\"\:\".*?\"',html)
        sls = re.findall(r'\"view_sales\"\:\".*?\"',html)
        for i in range(len(plt)):
            sales = eval(sls[i].split(':')[1])
            price = eval(plt[i].split(':')[1])
            title = eval(tlt[i].split(':')[1])
            ilt.append([sales , price , title])
    except:
        print("")
 
def printGoodsList(ilt):
    tplt = "{:4}\t{:8}\t{:16}\t{:32}"
    print(tplt.format("序号", "购买","价格", "商品名称"))
    count = 0
    for g in ilt:
        count = count + 1
        print(tplt.format(count, g[0], g[1],g[2]))

def main():
    goods = input('搜索商品:')
    depth = int(input('搜索页数:'))
    name = input('输入保存的excel名称(留空print):')
    start_url = 'https://s.taobao.com/search?q=' + goods
    infoList = []
    print('处理中...')
    for i in range(depth):
        try:
            url = start_url + '&s=' + str(44*i)
            html = getHTMLText(url)
            parsePage(infoList, html)
            print('第%i页成功...' %(i+1))
        except:
            continue
    writeExcel(infoList,name)
    print('完成!')

main()
