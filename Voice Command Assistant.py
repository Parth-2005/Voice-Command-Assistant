import pyttsx3  # pip install pyttsx3
import speech_recognition as sr  # pip install speechRecognition
import datetime
import wikipedia  # pip install wikipedia
import webbrowser
import os
import sys
import pyaudio
import random
import time
import subprocess
import ctypes
import win32com.client as wincl
import random
import pywhatkit
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
        engine.say(audio)
        engine.runAndWait()

def wishMe():
        hour = int(datetime.datetime.now().hour)
        if hour >= 0 and hour < 12:
            speak("Good Morning!")

        elif hour >= 12 and hour < 18:
            speak("Good Afternoon!")

        else:
            speak("Good Evening!")

        speak("How may I help you")

def takeCommand():
        # It takes microphone input from the user and returns string output

        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold=0.5
            audio = r.listen(source)

        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")

        except Exception as e:
            print(e)
            return "None"
        return query

browser_path="--Your browser's location--"
webbrowser.register('Chrome', None, webbrowser.Chrome(browser_path))
browser=webbrowser.get('Chrome')

if __name__ == "__main__":
        wishMe()
        while True:
            # if 1:
            query = takeCommand().lower()
            # Logic for executing tasks based on query
            if 'wikipedia' in query:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query)
                speak("According to Wikipedia")
                print(results)
                speak(results)

            elif 'youtube' in query:
                speak("opening youtube")
                query1 = query.replace("youtube ", "")
                browser.open(f"https://www.youtube.com/results?search_query={query1}")

            elif 'news' in query:
                speak("opening news")
                browser.open("https://www.news.google.com/")

            elif 'where is' in query:
                speak("opening maps")
                query1 = query.replace("where is ", "")
                browser.open(f"www.google.com/maps/place/{query1}")

            elif 'search' in query:
                query1 = query.replace("search ", "")
                speak("searching"+query1)
                browser.open(query1)

            elif 'who are you' in query:
                speak("i am your voice assistant")

            elif 'time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"Sir, the time is {strTime}")

            elif 'music' in query:
                music_dir = '--Your music directory--'
                songs = os.listdir(music_dir)
                print(songs)

                no=random.randint(0,85)
                os.startfile(os.path.join(music_dir, songs[no]))

            elif 'song' in query:
                query1 = query.replace(" song", "")
                speak("searching" + query1)
                browser.open("music.youtube.com/search?q="+query1)

            elif 'code' in query:
                speak("opening vscode")
                os.startfile('--VS code directory--')

            elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif 'goodbye' in query:
                sys.exit(speak("good bye"))

            elif 'shutdown' in query:
                speak('shutting down windows')
                pywhatkit.shutdown()