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

lst_news = ['bleacher-report', 'cnn', 'bbc-news', 'associated-press', 'the-verge', 'reuters','the-times-of-india','espn-cric-info','google-news','abc-news','business-insider','the-wall-street-journal','bbc-sport','bloomberg','buzzfeed','cbs-news','crypto-coins-news','engadget','entertainment-weekly','espn','fortune','four-four-two','fox-news','google-news-in','hacker-news','ign','medical-news-today','mtv-news','national-geographic','nbc-news','new-scientist','newsweek','next-big-future','nhl-news','polygon','recode','talksport','techcrunch']
print("News Channel:-\n")
print("(1) Bleacher-Report")
print("(2) CNN")
print("(3) BBC News")
print("(4) Associated-Press")
print("(5) The Verge")
print("(6) Reuters")
print("(7) The Times of India")
print("(8) ESPN Cricket Info")
print("(9) Google News")
print("(10) ABC News")
print("(11) Business Insider")
print("(12) The Wall Street Journal")
print("(13) BBC Sport")
print("(14) Bloomberg")
print("(15) Buzzfeed")
print("(16) CBS News")
print("(17) Crypto Coins News")
print("(18) Engadget")
print("(19) Entertainment Weekly")
print("(20) ESPN")
print("(21) Fortune")
print("(22) FourFourTwo")
print("(23) Fox News")
print("(24) Google News (India)")
print("(25) Hacker News")
print("(26) IGN")
print('(27) Medical News Today')
print('(28) MTV News')
print('(29) National Geographic')
print('(30) NBC News')
print('(31) New Scientist')
print('(32) NewsWeek')
print('(33) Next Big Future')
print('(34) NHL News')
print('(35) Polygon')
print('(36) Recode')
print('(37) TalkSport')
print('(38) TechCrunch')
 
choice = int(input("Which News Channel do you want to hear ? : "))

Speak(lst_news[choice-1])
 
 


 


