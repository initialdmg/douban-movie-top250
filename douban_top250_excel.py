#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

# 1 获取内容
baseuUrl = 'https://movie.douban.com/top250'

def fetch(start):
    # 25 per page
    url = baseuUrl + '?start=' + str(start)
    end = start + 25
    print('Fetch ', start, '-', end)
    html = requests.get(url)
    document = BeautifulSoup(html.content, features="html.parser")


    # 2 解析内容
    movies25 = document.select('ol.grid_view > li')

    for movie in movies25:

        title = movie.find('span', class_='title').get_text(strip=True) # 使用string 得到的类型是节点，不是python的string
        star = movie.find('span', class_='rating_num').get_text(strip=True)
        try:
            quote = movie.find('span', class_='inq').get_text(strip=True)
        except AttributeError as e:
            print('No Quote:', title)
            quote = ''

    # 3 保存
        # print(title, star, quote)
        cell = [title, star, quote]
        ws.append(cell)

    start = end
    if start == 250:
        return;
    else:
        #threading.Timer(3, fetch, (start,)).start()    # 会导致主线程直接运行完成，timer必须使用iterable传参数
        time.sleep(1)
        fetch(start)

print('Begin fetching data...')

# 1.1 新建excel
wb = Workbook()
# 1.2 创建工作表
ws = wb.active
# 1.3 初始化表头
ws.append(['影片名', '评分', '评语'])

fetch(0)

# Data can be assigned directly to cells
# ws['A1'] = 42

# Rows can also be appended
#ws.append([1, 2, 3])

print('Saving file...')

# Save the file
wb.save("douban_movie_top250.xlsx")

print('Done!')
