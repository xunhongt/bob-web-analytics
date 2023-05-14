#========================== Initialize Libraries =================================
# version:      1.0
# filename:     Collect top 200 Latest Threads in EDMW
# owner:        Bob

#================= Import Libraries and Initialize Functions =================

import requests
import time
import pandas

from bs4 import BeautifulSoup
from pandas import DataFrame

def getPage(url):
    req = requests.get(url)
    return BeautifulSoup(req.text, 'html.parser')

#================= Main Function =================

#------------- Collect Latest 200 Threads from EDMW ----------------

forumUrl = "https://forums.hardwarezone.com.sg"
forumSuffix = "/forums/eat-drink-man-woman.16"
pageCounter = 1

threadLinks = []
# Each Forum Page has 20 Threads, so range(11)

for i in range(1,3):
    pageUrl = forumUrl + forumSuffix + "/page-" + str(i)
    pageBsObj = getPage(pageUrl)

#-------------- Scrape EDMW Threads from the Forum and output into a list --------------
    for result in pageBsObj.find_all("a", {"data-preview-url":True}):
        threadLinks.append(result['href'])


#------------- Scrape Details from EDMW Threads ----------------
#  For each Thread, collect the following information: 
#
#    1. Thread Title 
#    2. Page Number 
#    3. Author of Thread Comment 
#    4. Thread Comment
#
#  Append these information into an array, which will be outputted into an excel file

titleArray = []
pageNoArray = []
authorArray = []
commentArray = []

#------------- Scrape Authors from EDMW Threads ----------------
#  For every new author, collect the following information: 
#
#    1. Author Name
#    2. First Joined
#    3. Message Count
#    4. Reaction Score

profileList = []

#------------ For each EDMW Thread collected, Scrape details from the Latest 10 pages ----------------

for thread in threadLinks:
    threadUrl = forumUrl + thread
    threadBsObj = getPage(threadUrl)

    #===== 1. Get the Pagination for each EDMW Thread (Max number of Pages)
    
        #----- if the thread has less than 10 pages, collect content from all pages within the thread   

    for result in threadBsObj.findAll('li', {"class": "pageNav-page"}):
        if len(result["class"]) == 1:
            pageCounter = int(result.text)
            break

    if pageCounter > 3:
        pageEnd = pageCounter - 3
    else:
        pageEnd = 0

        #----- Prepare URL for each Thread comments Page

    while pageCounter > pageEnd:
        threadPage = threadUrl + "page-" + str(pageCounter)

    #===== 2. Use BeautifulSoup to pull Thread Details

        threadBS = getPage(threadPage)
        
        #----- Get Thread Title

        threadTitle = threadBS.find('title').text.split(" | ")

        titleName = threadTitle[0]

        #----- Get Thread Page Number

        if "Page" not in threadTitle[1]:
            pageNumber = 1
        else:
            pageNumber = int((threadTitle[1].split(" "))[1])

        #----- Filter for all Comments within the Thread Page

        threadBody = threadBS.findAll('div', {'class', 'bbWrapper'})

        #----- Filter for all Authors within the Thread Page

        threadAuthors = threadBS.findAll('a', {'class':'username', 'itemprop':'name'})

        #----- Filter for Authors' details within the Thread Page  
      
        threadUserExtras = threadBS.findAll('div', {'class':'message-userExtras'})

    #===== 3. Collect Each Comment within a Thread Page

        authorCount = 1

        for comment in threadBody:
            if "Click to expand..." in comment.text:
                contentArray = comment.text.split("Click to expand...")
                commentArray.append(contentArray[1].replace("[\n\t]", " ").strip())
            else:
                commentArray.append(comment.text.replace("[\n\t]", " ").strip())

            print(threadPage)
            print(titleName)
            print(pageNumber)
            print(threadAuthors[authorCount].text)
            print(comment.text)

            print("\n")

            titleArray.append(titleName)
            pageNoArray.append(pageNumber)
            authorArray.append(threadAuthors[authorCount].text)

            authorCount = authorCount + 1

    #===== 4. Collect Author Details within a Thread Page

        authorCount = 1

        for profileDetails in threadUserExtras:
            authorName = threadAuthors[authorCount].text

            profileDetail = {
                "Name": authorName,
                "First Joined": profileDetails.text.strip().split("\n")[1],
                "Message Count": int(profileDetails.text.strip().split("\n")[5].replace(',', '')),
                "Reaction Score": int(profileDetails.text.strip().split("\n")[9].replace(',', ''))
            }
            profileList.append(profileDetail)

            authorCount = authorCount + 1


    #===== 5. Sleep for 3 seconds per page (to prevent website blocking)

        time.sleep(3)

        pageCounter = pageCounter - 1



profileList = list(map(dict, set(tuple(profile.items()) for profile in profileList)))


#------------- Output collected Results into an Excel ----------------

resultArray1 = DataFrame({'Title': titleArray, 'Page Number': pageNoArray, 'Author': authorArray, 'Comments': commentArray})
resultArray2 = DataFrame.from_records(profileList)

writer = pandas.ExcelWriter('edmw-comments-results-raw.xlsx')

resultArray1.to_excel(writer, sheet_name='raw_data', index=False)
resultArray2.to_excel(writer, sheet_name='profile_data', index=False)

writer.close()