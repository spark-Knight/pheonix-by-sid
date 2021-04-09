import subprocess
import wolframalpha
import pyttsx3
# import tkinter 
import json
import ecapture 
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
from webbrowser import Chrome
import os
import winshell
import pyjokes
# import feedparser
import smtplib 
import ctypes
import time
import requests 
from requests import get
import pyautogui
import selenium
import shutil
import pywhatkit as kit
import keyboard
import cv2
import phonenumbers
from ecapture import ecapture as ec
from phonenumbers import carrier
from phonenumbers import geocoder
# from twilio.rest import Client
# from clint.textui import progress
# from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty("rate",200)

f=open("settings.txt","r")
lines=f.readlines()
password=lines[0]
f.close()


Email={'hanish':'hanish.arora8@gmail.com','dinesh':'dkumar42358@gmail.com','sid':'dna8377850@gmail.com','nargis':'nannikhan72@gmail.com'}



def checker(key):
    if key in Email:
        return Email[key]
    else:
        return "no user found sorry"

url = 'https://www.google.com'
webbrowser.register('chrome',
	None,
	webbrowser.BackgroundBrowser("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))
# webbrowser.get('chrome').open(url)



def speak(audio):
    engine.say(audio)
    engine.runAndWait()
 
def wishMe():
    hour = int(datetime.datetime.now().hour)
    pt = time.strftime("%H:%M:%S")

    if hour>= 0 and hour<12:
        speak(f'Good Morning!,its {pt}')
        print(f'Good Morning!,its {pt}')
    elif hour>= 12 and hour<18:
        speak(f"Good Afternoon!,its {pt}")
        print(f"Good Afternoon!,its {pt}")    
    else:
        speak(f"Good Evening!,its {pt}") 
        print(f"Good Evening!,its {pt}")  
        speak(f"i am jarvis. i am your assistant. Please tell me how may i help you")

def usrname():
    speak("What should i call you")
    uname = takeCommand()
    speak("Welcome " + uname)
    columns = shutil.get_terminal_size().columns
     
    print("#####################".center(columns))
    print("Welcome!" , uname.center(columns))
    print("#####################".center(columns))
     
    speak("How can i Help you")

def takeCommand():
     
    r = sr.Recognizer()
     
    with sr.Microphone() as source:
         
        print("Listening...")
        r.pause_threshold = 0.5
        audio = r.listen(source)
  
    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")
  
    except Exception as e:
        print(e)    
        print("Unable to Recognize your voice.")  
        return "None"
     
    return query

def news():
    main_url = "https://newsapi.org/v2/top-headlines?sources=techcrunch&apikey=dd99e5301ef942e0a2748194efc4ad40"

    main_page = requests.get(main_url).json()
    articles = main_page["articles"]
    head=[]
    day=["first","second","third"]
    for ar in articles:
        head.append(ar["title"])
    for i in range(len(day)):
        speak(f"today's {day[i]} news is: {head[i]}")  
        print(f"today's {day[i]} news is: {head[i]}") 

def Music():
    speak("Tell me the name of song!")
    musicName = takeCommand() 
    kit.playonyt(musicName)

def closeApp():
    speak ("i have closed all programs!")  

    if "youtube" in query:
        os.system("TASKKILL /F /im chrome.exe ")

    elif "chrome" in query:
        os.system("TASKILL /F /im chrome.exe")

    elif "code" in query:
        os.system("TASKILL /F /im code.exe")

    elif "file explorer" in query:
        os.system("TASKILL /F / explorer.exe") 

    elif "cmd" in query:
        os.system("TASKILL /F /im cmd.exe") 



def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('siddhu4718k@gmail.com', password)
    server.sendmail('siddhu4718k@gmail.com', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
     
    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    usrname()
     
    while True:
         
        query = takeCommand().lower()
         
        # All the commands said by user will be 
        # stored here in 'query' and will be
        # converted to lower case for easily 
        # recognition of command

        if "trace my number" in query:
            speak("tell me your phone number")
            phone_number = input(int())
            ch_number = phonenumbers.parse(phone_number, "CH")
            speak("you exist in: " + (geocoder.description_for_number(ch_number, "en")))
            print(geocoder.description_for_number(ch_number, "en"))
            service_number = phonenumbers.parse(phone_number, "RO")
            speak("your sim name is: " + (carrier.name_for_number(service_number, "en")))
            print(carrier.name_for_number(service_number, "en"))


        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        # elif "wake up" in query:
        #       speak("tank's sir ")FF 
 
        # elif 'open youtube' in query:
        #     speak("Here you go to Youtube\n")
        #     webbrowser.Chrome.open("youtube.com")
        
        elif 'search in google' in query:
            speak("sir, what should i search in google")
            cm = takeCommand().lower()
            webbrowser.get("Chrome").open(f"{cm}")

        elif "open chrome" in query:
            os.startfile("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")

        elif "close all apps" in query:
            closeApp()    
        
        # elif "home page" in query:
        #     keyboard.press("Alt + Home")
        # elif "page up" in query:
        #     keyboard.press("ctrl + Page Down")
        # elif "page down" in query:
        #     keyboard.press("ctrl + page up")
        # elif "open incognito" in query:
        #     keyboard.press("alt + ctrl + n")
        # elif "downlaod" in query:
        #     keyboard.press("ctrl + j")
        # elif "open history" in query:
        #     keyboard.press("ctrl + h")
        # elif "forward a page" in query:
        #     keyboard.press("alt + right arrow") 
        # elif "new tab " in query :
        #     keyboard.press("ctrl + t")    
        # elif "close tab" in query :
        #     keyboard.press("ctrl + w")           


        elif "close chrome" in query:
            os.system("TASKKILL /F /im chrome.exe ") 

        
 
        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.get('chrome').open(url)

        elif "ip address" in query:
            ip = get("https://api.ipify.org").text
            speak(f"your ip address is {ip}")

 
        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.get("chrome").open("https://www.stackoverflow.com")   
 
        elif 'play music from my file' in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\91971\\OneDrive\\Desktop\\your dad\\songs"
            songs = os.listdir(music_dir)
            print(songs)
            rd = random.choice(songs)    
            random = os.startfile(os.path.join(music_dir, songs[0]))

        elif "play music from youtube" in query:
            Music()   

 
        elif "volume up" in query:
            pyautogui.press("volumeup")
        elif "volume down" in query:
            pyautogui.press("volumedown")
        elif "volume mute" in query:
            pyautogui.press("volumemute")           


        elif 'time' in query:
            strTime = datetime.datetime.now().strftime(" %H:%M:%S")    
            speak(f"Sir, the time is {strTime}")
            print(f"Sir, the time is {strTime}")
        # elif 'open opera' in query:
        #     codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
        #     os.startfile(codePath)



        elif 'send email' in query:
            try:
                speak("whom should i send mail to")                                                                                                                                                                                    
                to = checker(takeCommand().lower())
                speak("What should I say?")
                content = takeCommand()
                sendEmail(to, content)
                speak("Email has been sent !")
                print("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

    
 
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("what's about you, Sir")
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
 
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
 
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)
 
        elif 'exit' in query:
            speak("Thanks for using me sir, have a good day.")
            exit()
 
        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Sid.")
             
        elif 'joke' in query:
            speak(pyjokes.get_joke())
            print(pyjokes.get_joke())
             
        elif "calculate" in query:
            try:
                app_id = "R9V425-GXQQLELXJH"
                client = wolframalpha.Client(app_id)
                indx = query.lower().split().index('calculate') 
                query = query.split()[indx + 1:] 
                res = client.query(' '.join(query)) 
                answer = next(res.results).text
                print("The answer is " + answer) 
                speak("The answer is " + answer) 
            except Exception as e:
                speak("sorry sir i can't understand you, what you said")
                pass   
                 

        elif 'search'  in query:
            query = query.replace("search", "")
            webbrowser.open_new_tab(query)
            time.sleep(0)	
    
 
        elif "who i am" in query:
            speak("If you talk then definately you are human.")
 
        elif "why you came to world" in query:
            speak("Thanks to Sid. further It's a secret")
 
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant created by Sid")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Sid")
 
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                       0, 
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed succesfully")
 
        elif 'news' in query:
            speak("today news are")
            news()
         
        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
        elif 'shutdown system' in query or 'computer band karo' in query or "it's timing to leave "in query or "computer band kar" in query:
            speak("thank's sir ,for giving me your time.")
            # speak("Hold On a Sec ! Your system is on its way to shut down")
            speak("one minute sir!. let me close all background apps.")
            closeApp()
            os.system('shutdown /s /t 1')
            
                 
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
 
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
 
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.get("Chrome").open("https://www.google.co.in/maps/place/"  + location)
 
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")
 
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
 
        elif "write a note" in query:        
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "not dikhao" in query or "show note please" in query or "daayan not dikha" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(1))
 
        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)
             
            with open("Voice.py", "wb") as Pypdf:
                 
                total_length = int(r.headers.get('content-length'))
                 
                for ch in progress.bar(r.iter_content(chunk_size = 2391975),
                                       expected_size =(total_length / 1024) + 1):
                    if ch:
                      Pypdf.write(ch)
                     
        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "jarvis" in query:
             
            # wishMe()
            speak("yes sir!!!")
            # speak(assname)



        elif "weather" in query:
            speak("city name:")
            print("city name:")
            CITY = takeCommand()
            BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
            API_KEY = "1d7b48f22fc1a802e2be1716972facae"
            URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY
            response = requests.get(URL)
            if response.status_code == 200:
                data = response.json()
                main = data['main']
                temperature = main['temp']
                humidity = main['humidity']
                pressure = main['pressure']
                report = data['weather']
                print(f"{CITY:-^30}")
                speak(f"{CITY:-^30}")
                print(f"Temperature: {temperature}")
                speak(f"Temperature: {temperature}")
                print(f"Humidity: {humidity}")
                speak(f"Humidity: {humidity}")
                print(f"Pressure: {pressure}")
                speak(f"Pressure: {pressure}")
                print(f"Weather Report: {report[0]['description']}")
                speak(f"Weather Report: {report[0]['description']}")
            else:
                print("sorry sir i can't find the weather report")
                speak("sorry sir i can't find the weather report")

 
        elif "wikipedia" in query:
            webbrowser.get("Chrome").open("wikipedia.com")
 
        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assname)
 
        elif "will you be my gf" in query or "will you be my bf" in query:   
            speak("I'm not sure about, may be you should give me some time")
 
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
 
        elif "i love you" in query:
            speak("It's hard to understand")
 
        elif "what is" in query or "who is" in query:
             
            # Use the same API key 
            # that we have generated earlier
            client = wolframalpha.Client("R9V425-GXQQLELXJH")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")

        elif "youtube" in query or "let's enjoy on youtube " in query or "it's entertain timing" in query :
            speak("i also thinks that ,you are very tired, So let's do little bit enjoy sir...")
            speak("So, what video you want to play!")
            kit.playonyt(takeCommand()) 

        elif "close youtube" in query:
            speak()
            os.sysytem("TASKKILL /F /im chrome.exe ")

        elif "pause" in query:
            keyboard.press("space bar")
        elif "play again" in query:
            speak("sure, Sir!")
            keyboard.press("0")
        elif "mute video" in query:
            keyboard.press("m")
        elif "unmunte" in query:
            keyboard.press("m")    
        elif "next " in query:
            keyboard.press("shift + n") 
        elif "previous video" in query:
            keyboard.press("shift + p")    
        elif "skip" in query:
            keyboard.press("l")
        elif "back" in query:
            keyboard.press("j")   
        elif "full screen" in query:
            keyboard.press("f")   
        elif "theatre mode" in query:
            keyboard.press("t")    
        elif "miniplayer" in query:
            keyboard.press("i") 


        elif "where i am " in query or "where we are" in query or "location" in query:
            speak("wait sir, let me check")
            try:
                ipAdd = requests.get("https://api.ipify.org").text
                print(ipAdd)
                url = "https://get.geojs.io/v1/ip/geo/"+ ipAdd +".json"
                geo_requests = requests.get(url)
                geo_data = geo_requests.json()
                # state = geo_data["state"]
                city = geo_data["city"]
                country =geo_data["country"]
                speak(f"sir i am not sure, but i think we are in {city} city of {country} country ")
                print(f"sir i am not sure, but i think we are in {city} city of {country} country ")
            except Exception as e:
                speak("sorry sir, due to network issue i am not able to find where we are.")
                print("sorry sir, due to network issue i am not able to find where we are.")
                pass   

        elif "take screenshot" in query or "take a screenshot" in query:
            speak("sir, please tell me the name for this screenshot file")
            name = takeCommand().lower()
            speak("please sir hold the screen for few seconds, i am taking screenshot")
            time.sleep(1.5)
            img = pyautogui.screenshot()
            img.save(f"{name}.png")
            speak("i have done sir, the screenshot is saved in our main folder. Now i ready for next command")    
        

    



        
