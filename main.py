from newsapi import NewsApiClient
from win32com.client import Dispatch
 


def Speak(news_channel):
    speak = Dispatch("SAPI.Spvoice")
    api = NewsApiClient(api_key='a00d79a43a074feb8eba6ba88ecbef7d')

    # bbcnews = api.get_top_headlines(sources='bbc-news')
    news = api.get_top_headlines(sources=news_channel)
    news_articles = news['articles']
    # print(bbcnews)

    # print(len(bbc_news_1))

    author = f"From {news_channel}"
    (speak.Speak(author))


    for i in range(len(news_articles)):

        print(news_articles[i]['title'])
        speak.Speak(news_articles[i]['title'])
        print(news_articles[i]['description'])
        speak.Speak(news_articles[i]['description'])
        print('\n')

lst_news = ['bleacher-report', 'cnn', 'bbc-news', 'associated-press', 'the-verge', 'reuters']
print("News Channel:-\n")
print("(1) Bleacher-Report")
print("(2) CNN")
print("(3) BBC News")
print("(4) Associated-Press")
print("(5) The Verge")
print("(6) Reuters")
 
choice = int(input("Which News Channel do you want to hear ? : "))

Speak(lst_news[choice-1])
 
 


 


