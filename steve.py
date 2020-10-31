from tkinter import *
from PIL import Image, ImageTk
import pyttsx3
import subprocess
import wolframalpha
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import ctypes
import time
import requests
import shutil
import win32com.client as wincl
from urllib.request import urlopen
import pandas as pd
import numpy



file_csv = 'C:\\Users\\hp\\DATA SCIENCE\\Titanic_123.csv'
file_excel = 'C:\\Users\\hp\\DATA SCIENCE\\Untitled Folder 1\\BIT-6th.xlsx'


appid = 'REWA4P-4H4LLP5K63'


root = Tk()
root.title('steve')
root.geometry('600x800')


global var
global var2
global var3
global var4

var = StringVar()
var2 = StringVar()
var3 = StringVar()
var4 = StringVar()





def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour >= 0 and hour < 12:
        speak('Good Morning Sir!')

    elif hour >= 12 and hour < 18:
        speak('Good Afternoon Sir!')

    else:
        speak('Good evening Sir!')

    global assname

    assname = ('Steve')
    speak('i am your assistant')
    speak(assname)


def takeCommand():
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        print('Listening')
        listen = ('Listening...')
        r.pause_threshold = 1
        
        var2.set(listen)
        root.update()
        audio = r.listen(source)
        

    try:
        print('Recognizing...')
        recog = ('Recognizing...')
        # speak(recog)
        var2.set(recog)
        root.update()
        query = r.recognize_google(audio)
        var4 =(f'User said: {query}\n')
        answer.insert(END, var4)
        
        
    except Exception:
        
        print('say that again please')
        speak('say that again please')
        return 'None'
    
    return query

 

def uname():
    global u_name
    speak('what should i call you sir?')
    u_name = takeCommand()
    speak(u_name)
    var.set(u_name)
    root.update()
    greet = (f'Welcome Mr. {u_name}')
    speak(greet)



def Play():

    wishMe()
    uname()
    



    while True:
        query = takeCommand().lower()

        # question = takeCommand()
        # client = wolframalpha.Client(appid)
        # res = client.query(question)
        # res.results
        # var3 = next(res.results).text
        # answer.insert(END, var3)
        # speak(var3)
       

        if 'wikipedia' in query:
            speak(u_name)
            speak('Searching wikipedia')
            query =  query.replace('wikipedia','')
            results = wikipedia.summary(query, sentences = 2)
            var3 = (f'result:\n  {results}')
            answer.insert(END, var3)
            speak(results)

        elif 'change your name' in query:
            speak('what would you like to call me, Sir?')
            assname = takeCommand()
            mchn_name = (f'{assname}, Nice Name')
            var3 = mchn_name
            answer.insert(END, var3)
            speak(mchn_name)

        elif 'open youtube' in query:
            speak('youtube is opening sir')
            webbrowser.open('youtube.com')

        elif 'open google' in query:
            speak('google is opening sir')
            webbrowser.open('google.com')

        elif 'play music' in query:
            speak('ok sir, enjoy the music')
            music_dir = 'D:\songs'
            song = os.listdir(music_dir)
            var3 = song
            answer.insert(END, var3)
            os.startfile(os.path.join(music_dir, song[0]))

        elif 'whats the time' in query or 'the time' in query:
            strTime = datetime.datetime.now().strftime('%H: %M: %S')
            var3 = (f'Sir, the time is {strTime}')
            answer.insert(END, var3)
            speak(var3)

        elif 'tell me a joke' in query or 'joke' in query:
            var3 = pyjokes.get_joke()
            answer.insert(END, var3)
            speak(var3)

        elif 'search' in query or 'play' in query:
            query = query.replace('search', '')
            query=query.replace('play','')
            webbrowser.open(query)

        elif 'open new document' in query:
            speak('the WPS office is opening Sir')
            power = r'C:\\Users\\hp\AppData\\Local\\Kingsoft\\WPS Office\\ksolaunch.exe'
            os.startfile(power)

        elif 'change the background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,'D:\\wallpaper\\1109826.jpg',0)
            speak('backgorund changed successfully')

        elif 'lock window' in query:
            speak('locking the window')
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown the system' in query:
            speak('hold on a second, your system is on its way to shutdown')
            subprocess.call('shutdown /l')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False)
            speak('recycle bin recycled')

        elif "don't listen" in query or 'hold on' in query or 'wait' in query:
            speak('for how much time you want to stop javis from listening commands')
            a = int(takeCommand())
            var3 = (f'waiting {a}')
            answer.insert(END, var3)
            time.sleep(a)

        elif 'where is' in query:
            query = query.replace('where is','')
            location = query
            var3 = (f'you ask to locate {location}')
            answer.insert(END, var3)
            speak(var3)
            webbrowser.open('https://www.google.nl / maps / place/' + location + "")
        

        elif 'restart' in query:
            speak('restarting your system')
            subprocess.call(['shurdown', '/r'])

        # elif 'hibernating' or 'lock' in query:
        #     speak('hibernating')
        #     subprocess.call('shutdown / h')

        elif 'log off' in query or 'sign out' in query:
            speak('make sure all the application are closed before sign-out')
            time.sleep(5)
            subprocess.call(['shutdown', '/l'])

        elif 'activate engine' in query:
            speak('the wolframealpha engine has been activated')
            question = takeCommand()
            client = wolframalpha.Client(appid)
            res = client.query(question)
            res.results
            var3 = next(res.results).text
            answer.insert(END, var3)
            speak(var3)

        elif 'read csv data' in query:
            df = pd.read_csv(file_csv)
            speak('reading data')
            var3 = df.head()
            answer.insert(END, var3)

        elif 'shape the data' in query:
            df = pd.read_csv(file_csv)
            speak(f'ok {u_name} shaping data')
            var3 = df.shape
            answer.insert(END, var3)

        elif 'describe the data' in query:
            df = pd.read_csv(file_csv)
            speak(f'{u_name}, describing the data')
            var3 = df.describe()
            answer.insert(END, var3)

        elif 'show columns' in query:
            df = pd.read_csv(file_csv)
            speak(f'{u_name}, collecting columns')
            var3 = df.columns
            answer.insert(END, var3)

        elif 'sleep' in query or 'exit' in query:
            bye = (f"ok {u_name}, Good Bye")
            var3 = bye
            answer.insert(END, var3)
            speak(bye)
            exit()
        


   




imgframe = Frame(root)

path = 'zxcvjpg.jpg'
img = ImageTk.PhotoImage(Image.open(path))
panel = Label(root, image = img, height= 400, width = 600)
panel.pack(side='top', fill=X, ipadx = 30, pady = 15, padx = 15)
imgframe.pack()

# top frame 

topframe = Frame(root)
name = Label(topframe, textvariable = var, font = ("Courier", 16))
name.config(bg='#80b3ff')
name.pack(fill = X, ipadx = 150, pady = 15)
button = Button(topframe, text= 'play', width = 10, bg = '#ff704d', command= Play)
button.pack()

topframe.pack(side = TOP)

# Mid frame
midframe = Frame(root)

midframe.pack(fill = X, ipadx = 200, pady = 15)
label = Label(midframe, width=50, height = 2, font = ('Courier', 12), textvariable = var2)
label.pack()

# bottom frame
bottomframe = Frame(root)

scroll = Scrollbar(bottomframe)
scroll.pack(side=RIGHT, fill = Y)

answer = Text(bottomframe, width = 50, height = 15, yscrollcommand = scroll.set)
scroll.config(command = answer.yview)
# result = Label(answer, width = 50, height = 15, textvariable = var2)
# result.pack()
answer.pack()

bottomframe.pack(fill = X, ipadx = 200, pady = 15)


root.mainloop()