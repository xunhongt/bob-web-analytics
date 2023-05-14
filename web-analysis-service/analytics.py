###### Objective: A Web Scraping Program that performs sentiment analysis on the top 40 Latest Threads in EDMW ############

#================= Helper Function =================

import nltk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from nltk.tokenize import word_tokenize
from nltk.probability import FreqDist
from nltk.stem import WordNetLemmatizer
from nltk.corpus import words
from wordcloud import WordCloud, STOPWORDS
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

def getWordNorm(string):

# This function takes a word list and returns a list of normalized words
#   Input: string
#   Output: normalized string

    wordList = word_tokenize(string)
    english_vocab = set(word.lower() for word in words.words())
    wnl = nltk.WordNetLemmatizer()
    wordNorm = [wnl.lemmatize(word) for word in wordList if word in english_vocab]

    wordNorm = ' '.join(wordNorm)

    return wordNorm

def removeStopWords(string):

# This function removes stopwords from the string
#   Input: string
#   Output: string (without stop words)
    filterList = []
    wordList = word_tokenize(string)
    stopWords = STOPWORDS

    for word in wordList:
        if word not in stopWords:
            filterList.append(word)

    wordFilter = ' '.join(filterList)

    return wordFilter


def getFreqDistMostCommon(list, count):

# This function generates a frequency distribution from the list, and prints out the most common occurrences based on count
#   Input: wordlist, occurrences
#   Output: Common Occurrences in Freq Dist list based on input

    fdist = FreqDist(list)
    return fdist.most_common(count)

    
#================= Main Function =================

#------------- Clean Data ----------------

#==== Remove Duplicate Authors

#---- Sort Profile by Top Message Counts ----
profile_data = pd.read_excel("edmw-comments-results-raw.xlsx", sheet_name = 'profile_data')
profile_data = profile_data.sort_values(by=['Message Count'], ascending=False)

#---- Remove Duplicate Authors ----
profile_data = profile_data.drop_duplicates(subset=['Name'], keep="first")

#==== Remove NA Values from Thread Comments
content_data = pd.read_excel("edmw-comments-results-raw.xlsx", sheet_name = 'raw_data')
content_data = content_data.dropna()
content_data = content_data.drop_duplicates(subset=['Author','Comments'], keep="first")

#==== Save Cleaned Data
writer = pd.ExcelWriter("edmw-comments-results.xlsx")

content_data.to_excel(writer, sheet_name='raw_data', index=False)
profile_data.to_excel(writer, sheet_name='profile_data', index=False)

writer.close()

#------------- Analysis - Thread Titles ----------------

#---- Obtain the Thread Title information from Dataset ----

content_data = pd.read_excel("edmw-comments-results.xlsx", sheet_name = 'raw_data')
thread_title_data = content_data.drop_duplicates(subset=['Title'], keep="first")

thread_title_data = thread_title_data.sort_values(by=['Page Number'], ascending=False)

#---- Plot Top 10 Thread Titles against their Page Numbers ----

threadTitleArray = []
threadPageNo = []

for index, row in thread_title_data.head(10).iterrows():
    threadTitleArray.append(row["Title"])
    threadPageNo.append(row["Page Number"])

plt.rcParams['font.sans-serif']=['SimHei'] #Show Chinese label
plt.rcParams['axes.unicode_minus']=False
plt.rcParams.update({'font.size': 22})

plt.figure(figsize=(30,10))
threadTitles = plt.bar(threadTitleArray, threadPageNo, color = 'b', label = "Message Count")

plt.title('Top 10 Largest EDMW Threads')
plt.xlabel('Thread Titles')
plt.ylabel('Page Numbers')

plt.bar_label(threadTitles, padding=3)

#plt.subplots_adjust(bottom=0.100)  
plt.legend()
plt.xticks(rotation=90)
plt.show()

thread_title_wordPool = ''

for index, row in thread_title_data.iterrows():
    row["Title"] = row["Title"].lower()

    thread_title_wordPool = thread_title_wordPool + row["Title"] + " "

# Remove Stopwords from the Thread Title Word Pool
threadTitles = removeStopWords(thread_title_wordPool)

# Convert the collated string into a Token List
threadTitleTokens  = word_tokenize(threadTitles)
threadTitleTokens = [word for word in threadTitleTokens if word.isalnum()]

# Collate the top 50 most common words in the Token list
threadTitleCommon = getFreqDistMostCommon(threadTitleTokens, 50)


# Obtain the NLTK Text from Tokens
threadTitleText = nltk.Text(threadTitleTokens)


#------------- Produce Output  ----------------

print("================== Thread Titles - Frequency Distribution (Top 50 Words) ==============")
print("Top 50 most common words in Thread Titles: {}".format(threadTitleCommon))

print("\n")

print("List of Collocations (with window_size = 2): ")
threadTitleText.collocations(num=50, window_size=2)

print("\n")

#------------- Analysis of Authors ----------------

#---- Read Author Profile Data ----
profile_data = pd.read_excel("edmw-comments-results.xlsx", sheet_name = 'profile_data')
content_data = pd.read_excel("edmw-comments-results.xlsx", sheet_name = 'raw_data')

name = []
y1 = []
y2 = []

for index, row in profile_data.head(10).iterrows():
    name.append(row["Name"])
    y1.append(row["Message Count"])
    y2.append(row["Reaction Score"])    

name_loc = np.arange(len(name))
width = 0.25

plt.rcParams['font.sans-serif']=['SimHei'] #Show Chinese label
plt.rcParams['axes.unicode_minus']=False

plt.rcParams.update({'font.size': 22})

plt.figure(figsize=(30,10))
messageCount = plt.bar(name_loc - width/2, y1, width, color = 'b', label = "Message Count")
reactionScore = plt.bar(name_loc + width/2, y2, width, color = 'g', label = "Reaction Score")

plt.title('Top 10 Authors in EDMW by Message Count')
plt.xlabel('Authors')
plt.ylabel('Value')

plt.bar_label(messageCount, padding=3)
plt.bar_label(reactionScore, padding=3)

plt.legend()
plt.xticks(name_loc, name)
plt.show()


profile_data = profile_data.sort_values(by=['Reaction Score'], ascending=False)

name = []
y1 = []
y2 = []

for index, row in profile_data.head(10).iterrows():
    name.append(row["Name"])
    y1.append(row["Message Count"])
    y2.append(row["Reaction Score"])    

name_loc = np.arange(len(name))
width = 0.25

plt.rcParams['font.sans-serif']=['SimHei'] #Show Chinese label
plt.rcParams['axes.unicode_minus']=False

plt.rcParams.update({'font.size': 22})

plt.figure(figsize=(30,10))
messageCount = plt.bar(name_loc - width/2, y1, width, color = 'b', label = "Message Count")
reactionScore = plt.bar(name_loc + width/2, y2, width, color = 'g', label = "Reaction Score")

plt.title('Top 10 Authors in EDMW by Reaction Score')
plt.xlabel('Authors')
plt.ylabel('Value')

plt.bar_label(messageCount, padding=3)
plt.bar_label(reactionScore, padding=3)

#plt.subplots_adjust(bottom=0.100)
plt.legend()
plt.xticks(name_loc, name, rotation=90)
plt.show()

#---------- Generate Word Cloud for the top 10 Thread Authors by Message Count ------

profile_data = profile_data.sort_values(by=['Message Count'], ascending=False)

for profileIndex, profileRow in profile_data.head(10).iterrows():
	comment_words = ''	

	for contentIndex, contentRow in content_data.iterrows():
		if contentRow["Author"] == profileRow["Name"]:

			contentRow["Comments"] = contentRow["Comments"].lower()
			contentRow["Comments"] = contentRow["Comments"].replace(r'[-./?!,":;()\']',' ')
			

			comment_words = comment_words + contentRow["Comments"] + " "

	comment_words = removeStopWords(comment_words)
	#print(comment_words)

	wordcloud = WordCloud(width = 800, height = 800,
					background_color ='white',
					stopwords = STOPWORDS,
					min_font_size = 10).generate(comment_words)

	# plot the WordCloud image					
	plt.figure(figsize = (8, 8), facecolor = None)
	plt.imshow(wordcloud)
	plt.axis("off")
	plt.tight_layout(pad = 0)

	print(profileRow["Name"])
	plt.show()


#---------- Generate Word Cloud for the top 10 Thread Authors by Reaction Score ------

profile_data = profile_data.sort_values(by=['Reaction Score'], ascending=False)

for profileIndex, profileRow in profile_data.head(10).iterrows():
	comment_words = ''	

	for contentIndex, contentRow in content_data.iterrows():
		if contentRow["Author"] == profileRow["Name"]:

			contentRow["Comments"] = contentRow["Comments"].lower()
			contentRow["Comments"] = contentRow["Comments"].replace(r'[-./?!,":;()\']',' ')
			

			comment_words = comment_words + contentRow["Comments"] + " "

	comment_words = removeStopWords(comment_words)
	#print(comment_words)

	wordcloud = WordCloud(width = 800, height = 800,
					background_color ='white',
					stopwords = STOPWORDS,
					min_font_size = 10).generate(comment_words)

	# plot the WordCloud image					
	plt.figure(figsize = (8, 8), facecolor = None)
	plt.imshow(wordcloud)
	plt.axis("off")
	plt.tight_layout(pad = 0)

	print(profileRow["Name"])
	plt.show()


#------------- Text Processing for Thread Content  ----------------

# Separate the Comments based on Sentiment Compound value
# -- Remove escape characters (\n and \t)
# -- Normalize the Thread content 
# -- Separate the comment based on Sentiment Compound value
#      - if value >= 0.05, add into positive count
#      - if value <= -0.05, add into negative count
#      - else, add into Neutral count

sia = SentimentIntensityAnalyzer()

positiveCount = 0
negativeCount = 0
neutralCount = 0

for contentIndex, contentRow in content_data.iterrows():
    comment = contentRow["Comments"].replace("[\n\t]", " ")
    comment = getWordNorm(comment)
    siaCompound = sia.polarity_scores(comment)["compound"]
    if siaCompound >= 0.05:
        positiveCount = positiveCount + 1
    elif siaCompound <= -0.05:
        negativeCount = negativeCount + 1
    else:
        neutralCount = neutralCount + 1

#------------- Produce Output  ----------------

print("================== Comments - Sentiment Analysis ==============")
print("Positive Comments: {:.2f}%".format((positiveCount / len(content_data.index)) * 100))
print("Negative Comments: {:.2f}%".format((negativeCount / len(content_data.index)) * 100))
print("Neutral Comments: {:.2f}%".format((neutralCount / len(content_data.index)) * 100))

thread_comment_wordPool = ''

for index, row in content_data.iterrows():
    row["Comments"] = row["Comments"].replace("[\n\t]", " ")
    row["Comments"] = row["Comments"].lower()

    thread_comment_wordPool = thread_comment_wordPool + row["Comments"] + " "

# Remove Stopwords from the Thread Comment Word Pool
threadComments = removeStopWords(thread_comment_wordPool)

# Convert the collated string into a Token List
threadCommentsTokens  = word_tokenize(threadComments)
threadCommentsTokens = [word for word in threadCommentsTokens if word.isalnum()]

# Collate the top 50 most common words in the Token list
threadCommentsCommon = getFreqDistMostCommon(threadCommentsTokens, 50)


# Obtain the NLTK Text from Tokens
threadCommentsText = nltk.Text(threadCommentsTokens)


#------------- Produce Output  ----------------

print("================== Thread Comments - Frequency Distribution (Top 50 Words) ==============")
print("Top 50 most common words in Thread Comments: {}".format(threadCommentsCommon))

print("\n")

print("List of Collocations (with window_size = 2): ")
threadCommentsText.collocations(num=50, window_size=2)

print("\n")