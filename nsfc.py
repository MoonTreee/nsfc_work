import requests
import http.cookiejar as cookielib
import re
import os.path
from pytesseract import image_to_string
import pytesseract
from bs4 import BeautifulSoup
from PIL import Image

try:
    from PIL import Image
except:
    pass


tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata"'

# 构造Request headers
agent = 'Mozilla/5.0 (Windows NT 5.1; rv:33.0) Gecko/20100101 Firefox/33.0'
headers = {
    "Host": "npd.nsfc.gov.cn",
    "Referer": "http://npd.nsfc.gov.cn/fundingProjectSearchAction.action",
    'User-Agent': agent
}

# 使用cookie
session = requests.session()
session.cookies = cookielib.LWPCookieJar(filename="cookies")
try:
    session.cookies.load(ignore_discard=True)
except:
    print("Cookie未加载")


# 获取验证码
def get_captcha():
    captcha_url = "http://npd.nsfc.gov.cn/randomAction.action"
    r = session.get(captcha_url, headers=headers)
    with open("captcha.jpg", "wb") as f:
        f.write(r.content)
        f.close()
    # try:
    #     im = Image.open('captcha.jpg')
    #     im.show()
    #     im.close()
    # except:
    #     print(u'请到 %s 目录找到captcha.jpg 手动输入' % os.path.abspath('captcha.jpg'))
    captcha = image_to_string(Image.open('captcha.jpg'), config=tessdata_dir_config)
    return captcha


# 解析页面
def parse_content(content, key_word):
    soup = BeautifulSoup(content, "lxml")
    # lis = soup.find_all("li")
    lis = soup.select(".time_li")
    # print(lis)
    file_name = key_word+".txt"
    file = open(file_name, "w", encoding="utf8")
    for li in lis:
        for string in li.stripped_strings:
            line = string.replace("\t", "").replace("\n", "").replace("\r", "") + "\t"
            file.write(line)
            print(line, "")
        print("\n")
        file.write("\n")
    file.close()


def search(key_word):
    post_url = "http://npd.nsfc.gov.cn/fundingProjectSearchAction!search.action#"
    page_size, captcha = get_page_size(key_word)
    post_data= {
        "pageSize": page_size,
        "currentPage": 1,
        "fundingProject.keyword": key_word,
        "checkCode": captcha
    }
    page = session.post(post_url, data=post_data, headers=headers)
    # print(page.content.decode("utf8"))
    parse_content((page.content.decode("utf8")), key_word)
    session.cookies.save()
    to_excel(key_word)


def to_excel(file):
    path = file + ".txt"
    new_path = file + "_to_excel.txt"
    new_file = open(new_path, "w", encoding="utf8")
    with open(path, "r", encoding="utf8") as f:
        for line in f.readlines():
            # print(line)
            # 正则查找
            if re.search(r"(.*?)批准号", line):
                title = re.search(r"(.*?)批准号", line).group(1).replace("\t", "")
            # print(title)
            if re.search(r"批准号：(.*)项目类别", line):
                number = re.search(r"批准号：(.*)项目类别", line).group(1).replace("\t", "")
            # print(number)
            if re.search(r"项目类别：(.*)依托单位", line):
                class_ = re.search(r"项目类别：(.*)依托单位", line).group(1).replace("\t", "")
            if re.search(r"依托单位：(.*)项目负责人", line):
                school = re.search(r"依托单位：(.*)项目负责人", line).group(1).replace("\t", "")
            if re.search(r"项目负责人：(.*)资助经费", line):
                person_in_charge = re.search(r"项目负责人：(.*)资助经费", line).group(1).replace("\t", "")
            if re.search(r"资助经费：(.*)批准年度", line):
                fund = re.search(r"资助经费：(.*)批准年度", line).group(1).replace("\t", "")
            if re.search(r"批准年度：(.*)年", line):
                year = re.search(r"批准年度：(.*)年", line).group(1).replace("\t", "")
            # 结题项目的处理
            conclusion = re.search(r"结题项目：(.*)", line)
            if conclusion:
                conclusion = conclusion.group(1).replace("\t", "").replace("\n", "")
                keyword = re.search(r"关键词：(.*)结题项目", line).group(1).replace("\t", "").replace("\n", "")
                new_line = file + "\t" + title + "\t" + number + "\t" + class_ + '\t' + school + '\t' + person_in_charge + "\t" + fund + '\t' + year + '\t' + keyword + '\t' + conclusion
                new_file.writelines(new_line + '\n')
            else:
                keyword = re.search(r"关键词：(.*)", line).group(1).replace("\t", "").replace("\n", "")
                new_line = file + "\t" + title + "\t" + number + "\t" + class_ + '\t' + school + '\t' + person_in_charge + "\t" + fund + '\t' + year + '\t' + keyword
                new_file.writelines(new_line + '\n')
            print(new_line)
        f.close()


def get_page_size(key_word):
    post_url = "http://npd.nsfc.gov.cn/fundingProjectSearchAction!search.action#"
    captcha = get_captcha()
    post_data = {
        "pageSize": 10,
        "currentPage": 1,
        "fundingProject.keyword": key_word,
        "checkCode": captcha
    }
    page = session.post(post_url, data=post_data, headers=headers)
    # print(page.content.decode("utf8"))
    soup = BeautifulSoup(page.content.decode("utf8"), "lxml")
    page_size = soup.find("strong", style="color: red").string
    print(page_size)
    session.cookies.save()
    return int(page_size), captcha


if __name__ == "__main__":
    # search("文本挖掘")
    # get_page_size("文本挖掘")
    # print(image_to_string(Image.open('captcha.jpg'), config=tessdata_dir_config))
    f = open("key_words_list.txt", "r")
    key_words = f.readlines()
    print(key_words)
    for ky in key_words:
        search(ky.replace("\n", ""))
    f.close()


