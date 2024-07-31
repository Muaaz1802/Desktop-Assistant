import speech_recognition as sr
import win32com.client
# import openai
import webbrowser
import os
import subprocess
import datetime
import wikipedia

def greet():
    hour = datetime.datetime.now().hour
    if hour > 5 and hour < 12:
        say("Good Morning Sir.....")
    elif hour >= 12 and hour < 15:
        say("Good Afternoon Sir.....")
    else:
        say("Good Evening Sir.....")

def say(txt):
    speaker = win32com.client.Dispatch("SAPI.Spvoice")
    speaker.Speak(txt)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source: 
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            print("Recognising....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query

        except Exception as e:
            print("Some Error occured. Sorry from jarvis")
            return "Some Error occured. Sorry from jarvis"
    
if __name__ == "__main__":
    greet()
    say("I'm Jarvis A. I. How may I help you?")
    while True:
        print("Listening...")
        query = takeCommand().lower()

        if 'how are you' in query:
            print("I'M fine sir.")
            say("I'm fine sir.")

        sites = [["youtube", "https://youtube.com/"], ["wikipedia", "https://wikipedia.com/"], ["google", "https://google.com/"], ["whatsapp", "https://web.whatsapp.com/"], ["telegram", "https://web.telegram.org"], ["stack overflow", "https://stackoverflow.com/"], ["linked in", "https://www.linkedin.com/"], ["github", "https://github.com/"]]

        apps = [["code", "C:\\Users\\muaaz\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"], ["discord", "C:\\ProgramData\\muaaz\\Discord\\Update.exe --processStart Discord.exe"]]

        for app in apps:
            if f"open {app[0]}" in query:
                say(f"opening {app[0]}")
                subprocess.Popen(app[1])


        for site in sites:
            if f"open {site[0]}" in query:
                say(f"opening {site[0]} sir...")
                webbrowser.open(site[1])

        if 'play music' in query:
            say("opening music")
            # os.system("FILE_PATH")
            webbrowser.open("https://music.youtube.com/")

        if 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir, the time is {strTime}")
            say(f"Sir, the time is {strTime}")

        # os.system("cls")

        # else:
        #     print("Sorry, I didn't recognise you. Say it again please.")
        #     say("Sorry, I didn't recognise you. Say it again please.")