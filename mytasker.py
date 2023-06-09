import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os,wmi
import smtplib
import ctypes
import subprocess
import platform
import time
import datetime
import wikiquote
import win32api
import keyboard
import psutil
import win32clipboard
from time import sleep
from tkinter import *
from PIL import ImageTk
import random
import requests
import PyDictionary   
import pyautogui
import clipboard
from translate import Translator

data={
    'rashmi':'sharmarashmismlAgmail.com',
    'utpana':'utpanamarwah98@gmail.com',
    'aastha':'guptaaastha282002@gmail.com',
    'archana':'Archanasharma2001828@gmail.com',
    'gulshan':'Gucchii315@gmail.com',
    'sheetal':'Sheetalkumari662@gmail.com',
    'arsh':'arshika.gill123@gmail.com',
    'amit sir':'amit.luckyfriends@gmail.com',
    'gurpreet mam':'er.kaurgurpreet123@gmail.com',
    'satinder mam':'sk.gill78@gmail.com',
}

engine=pyttsx3.init('sapi5')         #speech api: help in synthesis and recognition of voice
voices=engine.getProperty('voices')    #getting details of current voice
engine.setProperty('voice',voices[1].id)

def start():
    while True :
        query = takecommand().lower()
        functionality(query)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()        #without this command,speech will not be audible to us

import cv2

# Function to capture a photo
def capture_photo():
    camera = cv2.VideoCapture(0)  # Use 0 for the default camera, or specify the camera index if multiple cameras are connected
    return_value, image = camera.read()
    if return_value:
        cv2.imwrite("photo.jpg", image)  # Save the captured photo as "photo.jpg"
        print("Photo captured!")
    camera.release()


def wishme():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Janny Mam Please tell me how may i help you?")

def takecommand():           #it takes microphone input from the user and returns string output
    r=sr.Recognizer()
    r.energy_threshold = 400
    with sr.Microphone() as source:
        print('Listening....')
        audio=r.listen(source)
    try:
        print("Recognizing....")
        query=r.recognize_google(audio, language='en-in')                   #using google for voice recognition
        print(f"User said: {query}\n")                                      #user query will be printed
    except Exception as e:
        print(e)
        print("Say that again Please....")
        return "None"
    return query

def sendEmail(to, content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('riyachoudhary8288@gmail.com','C:\\Users\\Amar Computer\\Documents\\pwd')
    server.sendmail('riyachoudhary8288@gmail.com',to,content)
    server.close()

def increase_brightness():
    brightness = win32api.EnumDisplaySettings(None, -1).Brightness
    win32api.EnumDisplaySettings(None, -1).Brightness = min(brightness + 10, 100)
    win32api.ChangeDisplaySettings(win32api.EnumDisplaySettings(None, -1), 0)
    print("Brightness increased by 10%.")

def secs2hours(secs):
    seconds = secs % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60
    return "{}hours {}minutes and {}seconds".format(hour, minutes, seconds)

def locklaptop():
    if platform.system()=='Windows':
        ctypes.windll.user32.LockWorkStation()
    elif platform.system() == 'Darwin':
        subprocess.call('"/System/Library/CoreServices/Menu Extras/User.menu/Contents/Resources/CGSession" -suspend')

def setalarm():
    speak("What time should I set the alarm for?")
    text = takecommand()
    try:
        timeobj = datetime.datetime.strptime(text, '%I:%M %p')  
        speak(f"Alarm set for {timeobj.strftime('%I:%M %p')}")
        while True:
            now = datetime.datetime.now()
            if now.hour == timeobj.hour and now.minute == timeobj.minute:
                speak("Wake up! It's time to start the day!")
                break
            time.sleep(60)  # check again in 60 seconds
    except:
        speak("Sorry, I couldn't set the alarm. Please try again.")


def functionality(query):
    wishme()
    while True:
        query=takecommand().lower()    #converting user query into lower case
        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia....')
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query, sentences=2)
            print("According to wikipedia")
            print(results)
            speak(results)
        
        elif 'thank you' in query:
            speak("Welcome mam")
        
        elif 'open youtube' in query:
            speak("Opening youtube for you, Mam")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Opening google for you, Mam")
            webbrowser.open("google.com")
            speak("Do you want me to search for you mam?")
            choice=takecommand().lower()
            if choice=='yes':
                speak("what do you want me to search mam?")
                search_item=takecommand().lower()
                speak("Searching for")
                speak(search_item)
                webbrowser.get('windows-default').open('https://google.com/search?q={}'.format(search_item))
            else:
                speak("Opening Google for you, mam")
                webbrowser.get('windows-default').open('http://www.google.com')
        
        elif 'click my photo' in query:
            # Main program
            speak("okay mam")
            capture_photo()
            
        elif 'open geeksforgeeks' in query:
            speak("Opening  for you, Mam")
            webbrowser.open("geeksforgeeks.org")

        elif 'open w3school' in query:
            speak("Opening for you, Mam")
            webbrowser.open("w3schools.com")

        elif 'play music' in query:
            music_dir="C:\\Users\\Amar Computer\\Music"   #add directory
            songs=os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'set alarm' in query:
            speak("Hello, how can I help you?")
            while True:
                text = takecommand().lower()
                if "set alarm" in query:
                    setalarm()
                elif "stop" in query:
                    speak("Goodbye!")
                    break
                else:
                    speak("Sorry, I didn't understand that.")

        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Mam the time is {strTime}")
        
        elif 'open visual studio' in query:
            speak("opening for you, Mam")
            codepath="C:\\Users\\Amar Computer\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe" #gives directory path
            os.startfile(codepath)

        elif 'open ms-word' in query:
            speak("opening for you, Mam")
            path="C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office Word 2007"
            os.startfile(path)

        elif 'open power-point' in query:
            speak("opening for you, Mam")
            path="C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office PowerPoint 2007"
            os.startfile(path)

        elif 'open excel' in query:
            speak("opening for you, Mam")
            path="C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office Excel 2007"
            os.startfile(path)

        elif 'open outlook' in query:
            speak("opening for you, Mam")
            path="C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office Outlook 2007"
            os.startfile(path)
        
        elif 'open settings' in query:
            speak("opening for you, Mam")
            os.system("ms-settings:")

        elif 'windowsupdate' in query:
            os.system("ms-settings:windowsupdate")

        elif 'display' in query:
            os.system("ms-settings:display")

        elif 'night light settings' in query:
            os.system("ms-settings:nightlight")

        elif 'advanced scaling settings' in query:
            os.system("ms-settings:display-advanced")
        
        elif 'connect to a wireless device' in query:
            os.system('ms-settings-connectabledevices:devicediscovery')

        elif 'graphics setting' in query:
            os.system("ms-settings:display-advancedgraphics")

        elif 'display orientation' in query:
            os.system("ms-settings:screenrotation")

        elif 'sound' in query:
            os.system("ms-settings:sound")

        elif 'manage sound devices' in query:
            os.system("ms-settings:sound-devices")
        
        elif 'app volume' in query:
            os.system("ms-settings:apps-volume")

        elif 'notifications and actions' in query:
            os.system("ms-settings:notifications")
        
        elif 'power and sleep' in query:
            os.system("ms-settings:powersleep")

        elif 'battery' in query:
            os.system("ms-settings:batterysaver")

        elif 'see which apps are affecting your battery life' in query:
            os.system("ms-settings:batterysaver-usagedetails")

        elif 'battery saver settings' in query:
            os.system("ms-settings:batterysaver-settings")

        elif 'storage' in query:
            os.system("ms-settings:storagesense")
        
        elif 'multitasking' in query:
            os.system("ms-settings:multitasking")

        elif 'about my pc' in query:
            os.system("ms-settings:about")

        elif 'show me bluetooth devices' in query:
            os.system("ms-settings:bluetooth")

        elif 'personalization' in query:
            os.system("stams-settings:personalization")

        elif 'date and time' in query:
            os.system("ms-settings:dateandtime")
        
        elif 'cortana' in query:
            os.system("ms-settings:cortana")

        elif 'talk to cortana' in query:
            os.system("ms-settings:cortana-talktocortana")

        elif 'open documents folder' in query:
            speak("Opening for you, Mam")
            documents_folder_path = os.path.expanduser("~/Documents")
            os.system(f"open {documents_folder_path}")

        elif 'open desktop folder' in query:
            speak("opening for you, Mam")
            desktop_folder_path = os.path.expanduser("~/Desktop")
            os.system(f"open {desktop_folder_path}")

        elif 'open downloads folder' in query:
            speak("opening for you, Mam")
            downloads_folder_path=os.path.expanduser("~/Downloads")
            os.system(f"open {downloads_folder_path}")

        elif 'open music folder' in query:
            speak("opening for you, Mam")
            music_folder_path=os.path.expanduser("~/Music")
            os.system(f"open {music_folder_path}")

        elif 'open pictures folder' in query:
            speak("opening for you, Mam")
            pictures_folder_path=os.path.expanduser("~/Pictures")
            os.system(f"open {pictures_folder_path}")
        
        elif 'open videos folder' in query:
            speak("opening for you, Mam")
            videos_folder_path=os.path.expanduser("~/Videos")
            os.system(f"open {videos_folder_path}")

        elif 'shutdown' in query:
            os.system('shutdown /s /t 1')

        elif 'restart' in query:
            os.system('shutdown /r /t 1')

        elif 'sleep' in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

        elif 'hibernate' in query:
            os.system("shutdown /h")

        elif 'who created you' in query:
            speak('My Creator is Riya')

        elif 'lock screen' in query:
            try:
                locklaptop()
            except sr.UnknownValueError:
                print("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                print("Could not request results from Google Speech Recognition service")

        elif 'brightness' in query:
            if 'decrease brightness' in query:
                speak("decreasing brightness")
                dec = wmi.WMI(namespace='wmi')
                methods = dec.WmiMonitorBrightnessMethods()[0]
                methods.WmiSetBrightness(30,0)
            elif 'increase brightness' in query:
                speak("increasing brightness")
                ins = wmi.WMI(namespace='wmi')
                methods = ins.WmiMonitorBrightnessMethods()[0]
                methods.WmiSetBrightness(100, 0)

        elif 'send email' in query:
            try:
                speak("What should i say?")
                content=takecommand()
                to=data
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry Mam. I am not able to send this email at the moment.")

        elif 'battery status' in query:
            battery = psutil.sensors_battery()
            plugged = battery.power_plugged
            percent = int(battery.percent)
            time_left = secs2hours(battery.secsleft)
            if plugged==True:
                speak("Don't worry mam, charger is plugged in")

            elif percent > 40:
                speak('mam, currently there is no need to charge me, i can survive for' + time_left)

            else:
                speak('mam, please connect the charger as soon as possible, i can survive only for' + time_left)
     
        elif 'what is antonym of' in query:
            replies = ['ok, let me check', 'ok, let me find', 'one moment please', 'one second please', 'just a moment',
                    'just a second', 'just a minute']
            speak(random.choice(replies))
            query = query.replace("what is antonym of", "")
            dictionary = PyDictionary()
            antonym_list = dictionary.antonym(query)
            speak('antonyms for {} are'.format(query))
            for i in range(len(antonym_list)):
                speak(antonym_list[i])

        elif 'motivate me' in query:
            list = wikiquote.quotes('Motivation')
            engine.setProperty('rate', 150)
            speak(random.choice(list))
            engine.setProperty('rate', 180)

        elif 'horoscope' and 'today' in query:
            speak('ok')
            speak('whats your sunsign, or zodiac sign?')
            sunsign=takecommand().lower()
            res = requests.get('https://horoscope-astrology.p.rapidapi.com/horoscope/{}'.format(sunsign))
            data = res.json()
            horoscope=data['horoscope']
            speak(horoscope)
            speak('thats all, mam')

        elif 'bye' in query:
            speak("Good Bye,Mam")
            root.destroy()
            exit()

        elif 'type for me' in query:
            speak('What do you want me to type')
            input_text=takecommand()
            keyboard.write(input_text)
            speak("typing done mam")

        elif 'change your gender' in query:
            speak("Why do you want me to change my gender")
            sleep(0.5)
            speak('Its not an easy task but still i will do it for you')
            gender=engine.getProperty('voice')
            if gender=='HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0':
                engine.setProperty('voice',voices[1].id)
            else:
                engine.setProperty('voice', voices[0].id)
            sleep(1)
            speak("Affermative")
            print(gender)

        elif "your capabilities" in query or "what can you do for me" in query:
            speak("I am having different capabilities")
            speak("I can search in wikipedia for you")
            speak("I can open youtube, google, stackoverflow")
            speak("I can play music for you, I can tell time")
            speak("I can also send emails for you")
            speak("I can open chrome and others softwares too, for you")
            speak("I will be happy to help you")
            speak("I am designed to talk formally with users, and very soon i will be having learning abilities")
            speak("well, My creator Riya has made me much smarter than other virtual assistant's")
            speak("currently, i am under development, and will have more features in future")
            speak("Thank you very much for asking")

        elif "who are you" in query or "tell me about yourself" in query or "introduce yourself" in query:
            speak("I am your MyTasker")
            speak("I am a creation of Riya")
            speak("I am currently under development and i am learning many more things")

        elif 'the time' in query:
            strtime = datetime.datetime.now().strftime("%H:%M:%S")
            speak("Mam the time is")
            speak(strtime)

        elif 'old are you' in query:
            speak("Age is just a number for me, but if you insist, let's say I'm perpetually adorable!")

        elif 'your name' in query:
            speak('myself mytasker, mam')

        elif 'who created you' in query:
            speak('My Creator is Riya')

        elif 'hello' and 'everyone' in query:
            speak('Hello Everyone! My self mytasker')
            speak("I hope you all are doing great")

        elif 'hello' in query:
            speak("Hello, mam")

        elif 'technology' in query:
            speak("I am created in Python programming language")

        elif "select all" in query:
            speak('selecting all')
            keyboard.press_and_release('ctrl+a')

        elif "copy" in query:
            speak("copying")
            keyboard.press_and_release('ctrl+c')
            speak('done')

        elif "undo" in query:
            speak("ok")
            keyboard.press_and_release('ctrl+z')

        elif "paste" in query:
            speak('pasting')
            keyboard.press_and_release('ctrl+v')

        elif 'read text' in query:
            speak("Ok mam, Copy any text to the clipboard, and i will read that text for you")
            keyboard.wait('ctrl+c')
            sleep(0.5)
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            speak(data)

        elif 'open stackoverflow' in query:
            speak("Opening StackOverflow for you, Mam")
            webbrowser.open('https://stackoverflow.com/')

        elif 'screenshot' in query:
            speak('Will you please provide a filename for it')
            filename=takecommand().lower()
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save(r'C:\\Users\\91987\\Desktop\\{}.png'.format(filename))
            speak('screenshot taken mam, check your desktop, i saved it there for you')
            speak("do you want me to show it")
            choice=takecommand().lower()
            if choice=='yes':
                speak('just a second, mam')
                os.startfile("C:\\Users\\91987\\Desktop\\{}.png".format(filename))
                sleep(1)
                speak("looks good")

        elif 'translate' in query:
            speak('What do you want me to translate')
            content = takecommand().lower()
            speak("In which language do you want to translate, mam")
            language=takecommand().lower()
            translator = Translator(to_lang=language)
            translation = translator.translate(content)
            clipboard.copy(translation)
            speak(translation)
            speak("I copied the translation to clipboard, for your convenience")

        elif 'open calculator' in query:
            speak('okay mam')
            os.system('calc')

        elif 'open task manager' in query:
            speak('okay mam')
            os.system('taskmgr.exe')

        elif 'open this pc' in query:
            speak('okay mam')
            os.system('explorer.exe')

        elif 'read text' in query:
            speak("Ok mam, Copy any text to the clipboard, and i will read that text for you")
            keyboard.wait('ctrl+c')
            sleep(0.5)
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            speak(data) 

        elif 'open notepad' in query:
            speak('Opening notepad')
            os.system('notepad')

        
        elif 'type for me' in query:
            speak('What do you want me to type')
            input_text=takecommand()
            keyboard.write(input_text)
            speak("typing done mam")
    

def buttonreleased(event):
    speak("Ready and listening mam")
    start()

wishme()
root = Tk()
root.configure(bg='white')
root.title("MyTasker Mic")
root.lift()
root.iconbitmap('C:\\Users\\Amar Computer\\Desktop\\Riya\\Major\\minijarvis2.ico')
photo = ImageTk.PhotoImage(file='C:\\Users\\Amar Computer\\Desktop\\Riya\\Major\\minijarvis1.png')
label = Label(root,image=photo, bg="white")
label.grid()
label.bind("<ButtonPress>", buttonreleased)
mainloop()





# import speech_recognition as sr
# import threading

# WAKE_WORD = "assistant"

# def handle_command(command):
#     # Implement your logic to handle the recognized command here
#     if "weather" in command:
#         print("Fetching weather information...")
#     elif "time" in command:
#         print("Fetching current time...")

# def listen_for_wake_word():
#     recognizer = sr.Recognizer()
#     microphone = sr.Microphone()

#     while True:
#         print("Listening for the wake word...")
#         with microphone as source:
#             recognizer.adjust_for_ambient_noise(source)
#             audio = recognizer.listen(source)

#         try:
#             text = recognizer.recognize_google(audio)
#             if WAKE_WORD in text.lower():
#                 print("Wake word detected. Listening for command...")
#                 with microphone as source:
#                     recognizer.adjust_for_ambient_noise(source)
#                     audio = recognizer.listen(source)

#                 command = recognizer.recognize_google(audio)
#                 print(f"Command: {command}")
#                 handle_command(command)
#         except sr.UnknownValueError:
#             print("Sorry, I couldn't understand the audio.")
#         except sr.RequestError as e:
#             print(f"Request error: {str(e)}")

# if __name__ == "__main__":
#     # Create and start the background listening thread
#     listening_thread = threading.Thread(target=listen_for_wake_word)
#     listening_thread.start()

#     # Perform any other initialization or setup here

#     # Your program will continue to run in the background,
#     # allowing the listening thread to handle wake word detection

#     # You can add additional code here to perform other tasks or actions while the virtual assistant is running

#     # Join the listening thread to the main thread to keep the program running indefinitely
#     listening_thread.join()
