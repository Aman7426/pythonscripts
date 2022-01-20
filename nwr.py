import requests
import json
import time
from win32com.client import Dispatch

def speak(s):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(s)

data = requests.get("https://newsapi.org/v2/everything?q=apple&from=2022-01-17&to=2022-01-17&sortBy=popularity&apiKey=a64be10fa41a42e19213b00e1e0c4a4b")

result = data.json()
print(result['status'])
# print(result)

news = result['articles']
print(news)

speak("Welcome to the CodeWithHarry News Channel")
speak("Here are the top ten news of the awesome country India")
speak("So our first news is ")
for i  in range(0,10):
    print(i)
    print(news[i]['description'])
    speak(news[i]['description'])
    if i>=9:
        break

    if i == 8:
        speak("So our last news for today is ")
    else:
        speak("Moving To Our next news")


speak("Thanks for listening ! Have a nice day")
speak("Keep coding")