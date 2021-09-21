import json
import requests
import win32com.client as wincl

def

speaker_number = 1
spk = wincl.Dispatch("SAPI.SpVoice")
vcs = spk.GetVoices()
SVSFlag = 11
spk.Voice
spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)




if __name__ == '__main__':

    spk.Speak("Hi Sir i am a news Assist ant so lets start news")

url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=91d1ed3c3bdf439b9044b51522feac0f"
news = requests.get(url).text
news_json = json.loads(news)
arts = news_json['articles']
i = 0
for article in arts:
    i = i+1
    print(f"{i} {article['title']}")
    spk.Speak(article['title'])

    spk.Speak("Moving to next news")




