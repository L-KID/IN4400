import queue
import time
import requests
import re
import random
import xlrd
import xlwt
import json
from bs4 import BeautifulSoup


def turnUrl(id):
    return "http://movie.douban.com/subject/" + str(id)


initial_id = 26683290
pattern1 = re.compile(r'subject/\d+')
pattern2 = re.compile('\d+')

#queue
url_queue = queue.Queue(maxsize=500)
url_queue.put(initial_id)

#set that store all movies that have been visited
candidates=set([initial_id])

##
workbook=xlwt.Workbook()
worksheet=workbook.add_sheet("sheet1")
inde = 0

while (len(candidates)<20 and not url_queue.empty()):
    # if url_queue.empty():
    #     print('empty')

    tempId = url_queue.get()

    # if this movie has been visited, skip to next round
    # if tempId in seen:
    #     continue

    time.sleep(1)

    tempUrl = turnUrl(tempId)
    print('AAAA',tempUrl)

    # get web page through http request
    r = requests.get(tempUrl)
    if r.status_code!=200:
        continue

    # parse html file
    soup = BeautifulSoup(r.text, "lxml")

    #######
    name=soup.find("span",attrs={"property":"v:itemreviewed"}).string
    ratings=soup.find("strong",attrs={"class":"ll rating_num"}).string
    noofvotes=soup.find("span",attrs={"property":"v:votes"}).string
    director=soup.find("a",rel="v:directedBy").string

    ##collect 4 actors
    actlist=[]
    for i in soup.find_all("a",rel="v:starring"):
        actlist.append(i.string)
        if len(actlist)>3:
            break

    typelist=[]
    for i in soup.find_all("span",attrs={"property":"v:genre"}):
        typelist.append(i.string)
        if len(typelist)>3:
            break

    relDate=soup.find("span",attrs={"property":"v:initialReleaseDate"}).string

    #information of one film
    info=[tempId,name,ratings,noofvotes,director,actlist,typelist,relDate]

    #write the information to the excel file
    for i in range(len(info)):
        print(str(i)+"aaa")
        worksheet.write(inde+1,i+1,info.pop())

    #if candidates is more than 300, stop collecting rec items
    if len(candidates)<300:
        recSect = soup.find("div", id='recommendations')
        if recSect == None:
            continue

        # collect all similiar movies
        for item in recSect.find_all("a"):
            if item.string == None:
                continue

            print("get item")

            temp = re.search(pattern2, re.search(pattern1, item["href"]).group()).group()

            #temp already visited?
            if temp not in candidates:
                url_queue.put(temp)
                candidates.add(temp)
                print(temp)


    # seen.add(tempId)
    # print("seen: ",len(seen))
    print("queue",url_queue.qsize())
    print("candidates",len(candidates))

    inde+=1

workbook.save('info_full.xls')
print("success")


###############get more information

# def write_row(rindex,candlist):
#     for i in range(len(candlist)):
#         # if len(i)<2:

