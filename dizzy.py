import pyttsx3            #python -m pip install pyttx3 to convert text to sppech recognition.
import speech_recognition as sr         #python -m pip install speechRecognition
import datetime                  #by default
import wikipedia                 #python -m pip install wikipedia
import webbrowser                #by default
import os                        #by default to perform windows operations                 #by default
import subprocess
import wolframalpha
import tkinter
import json
import random
import operator
import winshell
import pyjokes
import feedparser
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup               #for web scrapping
import win32com.client as wincl
from urllib.request import urlopen


chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

engine = pyttsx3.init('sapi5')            #sapi5 provides lsit of two voices
voices = engine.getProperty('voices')         

#print(voices[0].id)

engine.setProperty('voice', voices[1].id)            #set the voice id in which we want out programme to speak. it contains two voices [0] for Male and [1] for Female

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >=0 and hour <12:
        speak('good morning sir')
    elif hour >12 and hour <18:
        speak('good afternoon sir')
    else:
        speak('good evening sir')
    speak('I am Dizzy, Your voice assistance, speed 1 terahertz, memory 1 zigabyte , Please tell me how may I help you')

def takeCommand():
    #it takes microphone input from the user and return string output


    r = sr.Recognizer()
    with sr.Microphone() as source:

        r.adjust_for_ambient_noise(source, duration = 0.2)
        print('Listening...')
       # r.pause_threshold = 1.0
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio)
        print(f"User said:{query}\n")
    
    except Exception as e:
        #print(e)
 
        print("Say that again please...")
        return "None"
    return query
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('avii2000thakurr@gmail.com', 'Avincy2018')
    server.sendmail('avii2000thakurr@gmail.com', to, content)
    server.close()





if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()


        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif "open youtube" in query:
            webbrowser.open("youtube.com")
        elif "open google" in query:
            webbrowser.open("google.com")
        elif "open stackoverflow" in query:
            webbrowser.open("stackoverflow.com")
        elif "play music" in query:
            music_dir = 'C:\\Users\\abhi2\\Music'
            songs = os.listdir(music_dir)
            d=random.choice(songs)
            
            os.startfile(os.path.join(music_dir,d))
    


        elif "open photos" in query or "show me the picture" in query:
            photo_dir = 'C:\\Users\\abhi2\\OneDrive\\Pictures\\snaps'
            photos = os.listdir(photo_dir)
            print(photos)
            os.startfile(os.path.join(photo_dir,photos[0]))
        elif "the time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is{strTime}" )

        

        elif "open code " in query:
            codePath = "C:\\Users\\abhi2\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif "email to abhishek" in query:
            try:
                speak("what should i say?")
                content = takeCommand()
                to = "abhi2000thakurr@gmail.com"
                sendEmail(to, content)
                speak("email has been sent")
            except Exception as e:
                print(e)
                speak("sorry abhi, i am not able to send email")

        elif "quit" in query or "bye" in query:
            speak("quitting sir, thanks for your time!")
            exit()

        elif "power off" in query:
            speak("do u want to switch off the computer sir")
     
            # Input voice command
            take = takeCommand()
            choice = take
            if choice == 'yes':
                
                # Shutting down
                print("Shutting down the computer")
                speak("Shutting the computer")
                os.system("shutdown /s /t 30")
                
            if choice == 'no':
                
                # Idle
                print("Thank u sir")
                speak("Thank u sir")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "sleep" in query:
            speak("Sleeping")
            subprocess.call("shutdown / h")
        
        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif 'how are you' in query:
            speak("I am fine sir, Thanks for asking")
            speak("How are you, Sir")
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
 
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
 
        
         
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Abhishek.")


        elif 'joke' in query:
            speak(pyjokes.get_joke())
             
               
 
         
        elif "who i am" in query:
            speak("If you talk then definitely your human.")
 
        elif "why you came to world" in query:
            speak("Thanks to Gaurav. further It's a secret")
 
        elif 'power point' in query:
            speak("opening Power Point presentation")
            power = r"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(power)
 
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant created by Abhishek")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Abhishek ")
 

        elif 'open bluestack' in query:
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)        
 

       
                 
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
 
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Dizzy from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
 
        
 
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Dizzy Camera ", "img.jpg")
 
        
 
        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('dizzy.txt', 'w')
            
                     
        elif "show note" in query:
            speak("Showing Notes")
            file = open("dizzy.txt", "r")
            print(file.read())
            speak(file.read(6))
 
    
                     
        
        elif "Dizzy" in query:
             
            wishMe()
            speak("Dizzy 1 point o in your service Mister")     
                    
        elif "will you be my gf" in query or "will you be my bf" in query:  
            speak("I'm not sure about, may be you should give me some time")
 
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
 
        elif "i love you" in query:
            speak("It's hard to understand")
 
        elif "what is" in query or "who is" in query:
             
            # Use the same API key
            # that we have generated earlier
            client = wolframalpha.Client("API_ID")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")                 
            

                 # elif "" in query:
            # Command go here
            # For adding more commands
 
 


       

              
        