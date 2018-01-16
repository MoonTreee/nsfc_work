# nsfc_work
nsfc爬虫

本程序可针对http://npd.nsfc.gov.cn/ 中的项目信息进行自动化获取以及本地化保存。

# 1. 所需包
'''
import requests

import http.cookiejar as cookielib

import re

from pytesseract import image_to_string

from bs4 import BeautifulSoup
'''python
# 2. 使用方法

新建文件key_words_list.txt，将所需要检索的关键词写入到该文件中（按行写入），运行nsfc.py即可得到相关结果。


# 3. 迭代


a. 增加get_page_size()方法，自动获取检索结果数


b. 以列表形式传入key_word参数，支持一次获取并保存多个关键词的检索结果


c. 引入pytesseract模块，自动识别验证码


d. 合并原本的nsfc.py和to_excel.py,支持一步保存；
