import pandas as pd 
import datetime as datetime
from plyer import notification
import pyttsx3
import time
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices',voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def notification1(title,msg):
    notification.notify(
        title = title,
        message = msg,
        app_icon ="C:\\Users\\Harsh Sharma\\OneDrive\\Desktop\\Python Studies\\icons8-working-64.ico",
        timeout = 8
    )

df=pd.read_excel("C:\\Users\\Harsh Sharma\\OneDrive\\Desktop\\TimeTable\\tasks.xlsx")
current_date = datetime.datetime.now().date().strftime('%Y-%m-%d')
current_time = datetime.datetime.now().time().strftime('%H:%M')
for index,item in df.iterrows():
    d = item["Date"]
    t = item["Time"]
    if current_date == d  and current_time == t:
        a = item["Task"]
        x = int(item["Room"])
        notification1("Alert","Your Class of "+a+ f" is going to start in room number {x}")
        speak("Your Class of " +a+ f" is going to start in room number {x}")
        
        