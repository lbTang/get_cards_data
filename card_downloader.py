# -*- encoding: utf-8 -*-
'''
@File    :   card_downloader.py
@Time    :   2021/01/25 20:26:35
@Author  :   Tang Lianbin 
@Version :   1.0
@Desc    :   None
'''


import urllib.request
import io
import sys
import json
import urllib.parse
import bs4
import xlwt
from bs4 import BeautifulSoup
import sqlite3
import re

#定义输出结果的编码为utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')


class Robot(object):

    def work(self):
        url = r'http://ka.05321888.com/ka/taocan/index.html'
        html = self.get_decoded_html(url)
        cards_info = self.get_info(html)
        self.connect_sqlite(cards_info)


    def get_decoded_html(self,url):
        headers = {
        'Connection':' keep-alive',
        'Pragma':' no-cache',
        'Cache-Control':' no-cache',
        'Accept':'application/json, text/plain, */*',
        'User-Agent':' Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
        'DNT':' 1',
        'Accept-Language':' zh-CN,zh;q=0.9,zh-TW;q=0.8',
        'Cookie':'',
        'Accept-Encoding': ''
        }

        url = r'http://ka.05321888.com/ka/taocan/index.html'

        request = urllib.request.Request(url, headers = headers)
        response = urllib.request.urlopen(request)
        html = response.read()
        #print(html)
        html_decoded = html.decode('gbk','ignore')
        return html_decoded
    
    def get_info(self,html):
        soup = BeautifulSoup(html,"html.parser",from_encoding="gbk")
        tbody = soup.find_all('tbody')[0]

        result = []
        for tr in tbody.children:
            if isinstance(tr,bs4.element.Tag):
                #（1）网页链接
                page = tr['onclick'].split('\'')[1]
                detail_url = 'http://ka.05321888.com/ka/taocan/'+ page
                #（2）编号
                code = page.split('.')[0]
                #（3）卡名
                td = tr.find('td')
                card_name = td.text
                #（4）备注
                all_i = tr.find_all('i')
                addition = ''
                #（5）月租
                if (code!='0038'):
                    monthly_cost = re.search(r'(\d+元*)包', card_name, re.M|re.I).group(1)
                    if monthly_cost[-1] in '0123456789':
                        monthly_cost+='元'
                else:
                    monthly_cost = ''

                for i in all_i:
                    if(i.text != ""):
                        addition=addition+' '+i.text
                result.append({'code':code,'card_name':card_name,'addition':addition,'monthly_cost':monthly_cost,'detail_url':detail_url})
        return result
    

    def export_excel_from_db(self,write_info):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('sheet1')
        # 写入第一行内容  ws.write(a, b, c)  a：行，b：列，c：内容
        titleList = ['编号', '卡名', '附加', '月租', '套餐链接', '发布状态', '是否可申领']
        for i in range(0, len(titleList)):
            ws.write(0, i, titleList[i])

        # 所需获取数据对应key

        for row in range(0, len(write_info)):
            row_length = len(write_info[row])
            for column in range(0, row_length):
                # 文件中已写入一行title，所以这里写入内容时行号为row+1而非row
                print(row_length)
                if column < row_length-2:
                    ws.write(row + 1, column, write_info[row][column])
                elif column == row_length-2:
                    if write_info[row][column] == 0:
                        ws.write(row + 1, column, '未发布')
                    else:
                        ws.write(row + 1, column, '已发布')
                else:
                    if write_info[row][column] == 0:
                        ws.write(row + 1, column, '不可申领')
                    else:
                        ws.write(row + 1, column, '可以申领')
                    
                

        # 保存文件
        wb.save('./流量套餐列表_数据库.csv')
        wb.save('./流量套餐列表_数据库.xls')

    def connect_sqlite(self,cards_data):
        con = sqlite3.connect("cards_data.db")
        cur = con.cursor()
        # no:编号； card_name:卡名； addition:附加优惠信息； detail_url:套餐链接； ispublish:是否发布； state:生效状态
        cur.execute("CREATE TABLE IF NOT EXISTS cards(no, card_name, addition, monthly_cost, detail_url, ispublish, state)")
        fetch_cardNos = cur.execute("SELECT no FROM cards").fetchall()
        card_no_list = [item[0] for item in fetch_cardNos]
        for card in cards_data:
            if card['code'] not in card_no_list:
                cur.execute("""INSERT INTO cards VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (card["code"],card["card_name"],card["addition"],card["monthly_cost"],card["detail_url"],0,1)
                )
        con.commit()
        res = cur.execute("SELECT * FROM cards")
        print(res.fetchall())
        self.export_excel_from_db(res.fetchall())



if __name__ == '__main__':
    r1 = Robot()
    r1.work()
    
