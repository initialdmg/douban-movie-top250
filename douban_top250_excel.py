#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# 1 获取内容
url = 'https://movie.douban.com/top250'

print('Begin fetching data...')

html = requests.get(url)
document = BeautifulSoup(html.content, features="html.parser")

# 1.1 新建excel
wb = Workbook()
# 1.2 创建工作表
ws = wb.active
# 1.3 初始化表头
ws.append(['影片名', '评分', '评语'])

# 2 解析内容
movies50 = document.select('ol.grid_view > li')

for movie in movies50:

    title = movie.find('span', class_='title').get_text(strip=True) # 使用string 得到的类型是节点，不是python的string
    star = movie.find('span', class_='rating_num').get_text(strip=True)
    quote = movie.find('span', class_='inq').get_text(strip=True)

# 3 保存
    # print(title, star, quote)
    cell = [title, star, quote]
    ws.append(cell)

# Data can be assigned directly to cells
# ws['A1'] = 42

# Rows can also be appended
#ws.append([1, 2, 3])

print('Saving file...')

# Save the file
wb.save("douban_movie_top250.xlsx")

print('Done!')