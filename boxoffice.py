import requests
from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter

cookie = {"Cookie": "__utma=47724067.1298609274.1484304707.1484304707.1484316708.2; __utmb=47724067.0.10.1484316708; __utmc=47724067; __utmz=47724067.1484304707.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); Hm_lvt_e71d0b417f75981e161a94970becbb1b=1484304819,1484316708; Hm_lpvt_e71d0b417f75981e161a94970becbb1b=1484317661; time=MTEzNTI2LjIxNjM0Mi4xMDI4MTYuMTA3MTAwLjExMTM4NC4yMDc3NzQuMTE5OTUyLjExMTM4NC4xMDQ5NTguMTExMzg0LjExOTk1Mi4xMTEzODQuMTA5MjQyLjEwNDk1OC4xMTk5NTIuMTA5MjQyLjExMTM4NC4xMDQ5NTguMA%3D%3D; DIDA642a4585eb3d6e32fdaa37b44468fb6c=3tij5ppd3h6nr49q6osf0q6qe2; remember=0"}
row_num = 0
col_num = 0
x = 1
book = xlsxwriter.Workbook(r'data2012.xls')
tmp = book.add_worksheet()
page = 1
for l in range(1, 11):
    if l == 1:
        res = requests.get("http://58921.com/alltime/2012", cookies=cookie)
    else:
        res = requests.get("http://58921.com/alltime/2012?page=%s" % page, cookies=cookie)
        page += 1

    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text, "html.parser")
    tbody = soup.find("tbody")

    tr = tbody.tr
    td = tr.td
    print(td.get_text())
    tmp.write(row_num, col_num, td.get_text())
    col_num += 1

    for i in range(1, 7):
        td = td.find_next_sibling("td")
        if i == 3:
            img = td.img
            imgurl = img['src']
            print(imgurl)
            urllib.request.urlretrieve(imgurl, "image2012/%s.jpg" % x)
            x += 1
        else:
            print(td.get_text())
            tmp.write(row_num, col_num, td.get_text())
            col_num += 1
    row_num += 1
    col_num = 0

    for j in range(1, 20):
        tr = tr.find_next_sibling("tr")
        td = tr.td
        print(td.get_text())
        tmp.write(row_num, col_num, td.get_text())
        col_num += 1
        for k in range(1, 7):
            td = td.find_next_sibling("td")
            if k == 3:
                img = td.img
                imgurl = img['src']
                print(imgurl)
                urllib.request.urlretrieve(imgurl, "image2012/%s.jpg" % x)
                x += 1
            else:
                print(td.get_text())
                tmp.write(row_num, col_num, td.get_text())
                col_num += 1
        row_num += 1
        col_num = 0


