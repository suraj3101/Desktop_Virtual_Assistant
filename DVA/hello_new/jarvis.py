import random
import pyttsx3
import speech_recognition as sr
import datetime
import os
import time
import cv2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests import get
import wikipedia
import webbrowser
import pywhatkit
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import requests
from bs4 import BeautifulSoup
import speedtest
from playsound import playsound
import subprocess
import wolframalpha
import tkinter
import json
import random
import operator
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import shutil
from twilio.rest import Client
# from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import imdb
import pygame
import threading
from tkinter import *
from PIL import ImageTk

# Start
search = pyttsx3.init()


# Speak Function
def speak(audio):
    print(audio)
    voices = search.getProperty("voices")
    search.setProperty("voice", voices[3].id)
    search.setProperty("rate", 185)
    search.say(audio)
    search.runAndWait()


# Take Command from User:
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 2
        # audio = r.listen(source, timeout=1, phrase_time_limit=5)
        audio = r.listen(source, phrase_time_limit=5)

    try:
        print("Recognizing..")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        speak("Unable to Recognize..")
        return "None"
    return query


# Wish in initial phase:
def wish():
    hour = int(datetime.datetime.now().hour)

    if hour >= 0 and hour < 12:
        speak("good morning")
    elif hour >= 12 and hour < 18:
        speak("good afternoon")
    else:
        speak("good evening")

    assname = ("Jarvis 2 point o")
    speak("i am your assistant")
    speak(assname)


def username():
    speak("What should i call you sir")
    uname = takecommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))
    speak("How can i Help you, Sir")


def quitApp():
    hour = int(datetime.datetime.now().hour)
    if hour >= 3 and hour < 18:
        print("have a good day sir")
        speak("have a good day sir")
    else:
        print("Goodnight sir")
        speak("Goodnight sir")
        print("Offline")
        exit(0)


# done-1:
def open_notepad():
    npath = "C:\\Windows\\notepad.exe"
    os.startfile(npath)

# done-2:


def find_ip():
    ip = get('https://api.ipify.org').text
    speak(f"your IP address is {ip}")


# done-3:
def open_ar():
    adobe_path = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Adobe Reader X.lnk"
    os.startfile(adobe_path)


# done-4:
def open_cmd():
    cmd_path = "C:\\Users\\HP\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\System " \
        "Tools\\Command Prompt.lnk"
    os.startfile(cmd_path)


# done-5:
def find_temp():
    speak("Tell me the city!")
    city = takecommand()
    url = f"https://www.google.com/search?q=temperature+in+{city}"
    r = requests.get(url)
    data = BeautifulSoup(r.text, "html.parser")
    temperature = data.find("div", class_="BNeawe iBp4i AP7Wnd").text
    speak(f"The Temperature in {city} is: {temperature}")

# done-6:


def findOnGoogle(query):
    query = query.lower()
    query = query.replace("jarvis", "")
    query = query.replace("on google", "")
    query = query.replace("google", "")
    query = query.replace("what do you mean by", "")
    query = query.replace("define", "")
    query = query.replace("what is the definition of", "")
    # query = query.replace("who is", "")
    query = query.replace("according to", "")
    speak("This is what i found on the web")
    pywhatkit.search(query)

    try:
        import wikipedia as googleScrap
        result = googleScrap.summary(query, 3)
        speak(result)
    except:
        speak(
            "I'm sorry, I didn't understand that. Let me search the web for you.")
        webbrowser.open_new_tab(
            "https://www.google.com/search?q=" + query)


# done-7:
def timeToWakeUp():
    speak("Enter the time: ")
    # Time should be in 24hrs format:
    # E.g: 19:35:00
    time = input(": Enter the Time: ")

    pygame.init()
    pygame.mixer.init()

    while True:
        Time_Ac = datetime.datetime.now()
        now = Time_Ac.strftime("%H:%M:%S")

        if now == time:
            speak("Time to Wake Up!")
            pygame.mixer.music.load('alarm.mp3')
            pygame.mixer.music.play()
            # Wait for the sound to finish playing
            while pygame.mixer.music.get_busy():
                pygame.time.Clock().tick(10)  # Adjust tick rate if necessary
            speak("Alarm Closed")
            break

        elif now > time:
            break


# done-8:
def resultOnWikipedia(query):
    query = query.lower()
    speak("Searching on Wikipedia...")
    query = query.replace("wikipedia", "")
    query = query.replace("can you tell me", "")
    query = query.replace("let's talk about", "")
    query = query.replace("okay, let's talk about", "")
    query = query.replace("okay, so let's talk about", "")
    query = query.replace("talk about", "")
    query = query.replace("more about", "")
    query = query.replace("more about", "")
    results = wikipedia.summary(query, sentences=2)
    speak("according to wikipedia ")
    speak(results)
    print(results)

# done-9:


def findCurrentTime():
    strTime = datetime.datetime.now().strftime("%H:%M:%S")
    speak(f"The time is: {strTime}")

# done-10:


def changeName(query):
    query = query.replace("change my name to", "")
    assname = query
    speak(f"Your name is: {assname}")

# done-11:


def tellJoke():
    speak(pyjokes.get_joke())

# done


def tellNews():
    try:

        jsonObj = urlopen(
            '''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
        data = json.load(jsonObj)
        i = 1
        speak('here are some top news from the times of india')
        print('''=============== TIMES OF INDIA ============''' + '\n')
        for item in data['articles']:
            print(str(i) + '. ' + item['title'] + '\n')
            print(item['description'] + '\n')
            speak(str(i) + '. ' + item['title'] + '\n')
            i += 1
    except Exception as e:
        print(str(e))

# done-13:


def lockWindow():
    speak("locking the device")
    ctypes.windll.user32.LockWorkStation()


def shutDown():
    speak("Hold On a Sec ! Your system is on its way to shut down")
    subprocess.call('shutdown / p /f')

# done-15:


def emptyBin():
    winshell.recycle_bin().empty(confirm=False, show_progress=True, sound=True)
    speak("Recycle Bin Recycled")

# done-16:


def StopListen():
    # speak("for how much time you want to stop jarvis from listening commands")
    # a = int(takecommand())
    # time.sleep(a)
    # print(a)
    speak("For how much time do you want to stop Jarvis from listening to commands?")
    speak("enter the time in seconds")
    num = input(": Enter the Time: ")
    stop_time = int(num)
    speak(f"Stop listening for {stop_time} ")
    time.sleep(stop_time)
    print(stop_time)


# done-17:
def locationOnMap(query):
    query = query.replace("where is", "")
    query = query.replace("on map", "")
    query = query.replace("locate the city", "")
    query = query.replace("location of", "")
    location = query
    speak("User asked to Locate")
    speak(location)
    webbrowser.open("https://www.google.nl/maps/place/" + location + "")


# done-18:
def capturePhoto():
    ec.capture(0, "Jarvis Camera ", "img.jpg")

# done-19:


def restartSystem():
    subprocess.call(["shutdown", "/r"])

# done-20:


def writeNote():
    speak("What should I write, sir?")
    note = takecommand()
    file = open('jarvis.txt', 'r+')  # Open file in read/write mode
    speak("Sir, should I include date and time?")

    while True:
        snfm = takecommand()
        if 'yes' in snfm or 'sure' in snfm:
            now = datetime.datetime.now()
            file.write(f"\n\nDate: {now.strftime('%Y-%m-%d')}\n")  # Add date
            file.write(f"Time: {now.strftime('%H:%M:%S')}\n")  # Add time
            file.write(f"Note: {note}\n")
            break  # Exit loop if recognized successfully
        elif 'no' in snfm:
            file.write(f"Note: {note}\n")
            break  # Exit loop if recognized successfully
        else:
            speak("Sorry, I couldn't understand. Please say yes or no.")

    file.close()
    speak("Note added successfully, sir.")


# done-21:

def showNote():
    speak("Showing Notes")
    file = open("jarvis.txt", "r")
    notes = file.readlines()
    file.close()

    # Extracting the most recent note
    # Extract the last 3 lines (date, time, note)
    most_recent_note = notes[-3:]
    recent_date = most_recent_note[0].split(": ")[1].strip()
    recent_time = most_recent_note[1].split(": ")[1].strip()
    recent_note = most_recent_note[2].split(": ")[1].strip()

    # Speaking the most recent note's details
    speak(
        f"The most recent note is created at {recent_time} on {recent_date}, and the content is: {recent_note}")

    file.close()


# done-


def detect_image():
    img = cv2.imread('1.jpg')
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    haar_cascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
    faces_rect = haar_cascade.detectMultiScale(gray_img, 1.1, 9)
    for (x, y, w, h) in faces_rect:
        cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)
        cv2.imshow('Detected faces', img)
        cv2.waitKey(0)


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()

# done


def findMovies():
    moviesdb = imdb.IMDb()
    speak("Tell me the complete name of the movie: ")
    text = takecommand()
    movies = moviesdb.search_movie(text)
    # speak("Searching for " + text)
    speak(f'Searching for movie name as: {text}')
    if len(movies) == 0:
        speak("No result found")
    else:
        speak("I found these:")
        for movie in movies:
            title = movie['title']
            year = movie['year']
            speak(f'{title}-{year}')
            info = movie.getID()
            movie = moviesdb.get_movie(info)
            title = movie['title']
            year = movie['year']
            rating = movie['rating']
            plot = movie['plot outline']
            if year < int(datetime.datetime.now().strftime("%Y")):
                speak(
                    f'{title}was released in {year} has IMDB rating of {rating}.\nThe plot summary of movie is {plot}')
                print(
                    f'{title}was released in {year} has IMDB rating of {rating}.\nThe plot summary of movie is {plot}')
                break
            else:
                speak(
                    f'{title}will release in {year} has IMDB rating of {rating}.\The plot summary of movie is {plot}')
                print(
                    f'{title}will release in {year} has IMDB rating of {rating}.\The plot summary of movie is {plot}')
                break

# done


def open_application(input):

    if "chrome" in input:
        speak("Opening Google Chrome")
        os.startfile(
            "C:\Program Files\Google\Chrome\Application\chrome.exe")
        return

    elif "firefox" in input or "mozilla" in input or "mozilla firefox" in input:
        speak("Opening Mozilla Firefox")
        os.startfile(
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Firefox Developer Edition.lnk")
        return

    elif "word" in input:
        speak("Opening Microsoft Word")
        os.startfile(
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")
        return

    elif "excel" in input:
        speak("Opening Microsoft Excel")
        os.startfile(
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
        return

    elif "powerpoint" in input:
        speak("Opening Microsoft Power Point")
        os.startfile(
            "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\PowerPoint.lnk")

    elif "onenote" in input:
        speak("Opening Microsoft OneNote")
        os.startfile(
            "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\OneNote.lnk")

    elif "microsoft edge" in input or "edge" in input:
        speak("Opening Microsoft Edge")
        os.startfile(
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")

    elif "Brave" in input:
        speak("Opening Brave")
        os.startfile(
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Brave.lnk")

    elif "adobe reader" in input:
        speak("Opening Adobe Reader")
        os.startfile(
            "C:\\ProgramData\\Microsoft\Windows\Start Menu\Programs\Adobe Reader X.lnk")

    elif "notepad" in input:
        speak("Opening Notepad")
        os.startfile(
            "C:\\Program Files\\WindowsApps\\Microsoft.WindowsNotepad_11.2402.22.0_x64__8wekyb3d8bbwe\\Notepad\\Notepad.exe")

    elif "cmd" in input or "command prompt" in input:
        speak("Opening cmd")
        os.startfile(
            "C:\\Users\\HP\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\System Tools\\Command Prompt.lnk")

    else:
        input = input.replace(" ", "")
        webbrowser.open(f"http://www.{input}.com")
        return


# done
def close_application(application_name):

    if 'chrome' in application_name.lower():
        subprocess.call("TASKKILL /F /IM chrome.exe", shell=True)

    elif 'firefox' in application_name.lower() or 'mozilla' in application_name.lower():
        subprocess.call("TASKKILL /F /IM firefox.exe", shell=True)

    elif 'word' in application_name.lower():
        subprocess.call("TASKKILL /F /IM winword.exe", shell=True)

    elif 'excel' in application_name.lower():
        subprocess.call("TASKKILL /F /IM excel.exe", shell=True)

    elif 'powerpoint' in application_name.lower():
        subprocess.call("TASKKILL /F /IM powerpnt.exe", shell=True)

    elif 'onenote' in application_name.lower():
        subprocess.call("TASKKILL /F /IM onenote.exe", shell=True)

    elif 'edge' in application_name.lower():
        subprocess.call("TASKKILL /F /IM msedge.exe", shell=True)

    elif 'brave' in application_name.lower():
        subprocess.call("TASKKILL /F /IM brave.exe", shell=True)

    elif 'adobe reader' in application_name.lower():
        subprocess.call("TASKKILL /F /IM AcroRd32.exe", shell=True)

    elif 'notepad' in application_name.lower():
        subprocess.call("TASKKILL /F /IM notepad.exe", shell=True)

    elif 'cmd' in application_name.lower() or 'command prompt' in application_name.lower():
        subprocess.call("TASKKILL /F /IM cmd.exe", shell=True)

    else:
        print(f"Unable to close {application_name}.")

    # Dictionary mapping application names to their corresponding process names
    # app_processes = {
    #     "chrome": "chrome.exe",
    #     "firefox": "firefox.exe",
    #     "word": "winword.exe",
    #     "excel": "excel.exe",
    #     "powerpoint": "powerpnt.exe",
    #     "onenote": "onenote.exe",
    #     "edge": "msedge.exe",
    #     "brave": "brave.exe",
    #     "adobe reader": "AcroRd32.exe",
    #     "notepad": "notepad.exe",
    #     "command prompt": "cmd.exe"
    # }

    # if application_name is None:
    #     application_name = input(
    #         "Enter the name of the application to close: ")

    # # Convert the application name to lowercase for case-insensitive comparison
    # application_name_lower = application_name.lower()

    # # Check if the lowercase application name is in the dictionary
    # if application_name_lower in app_processes:
    #     process_name = app_processes[application_name_lower]
    #     try:
    #         os.system(f"TASKKILL /F /IM {process_name}")
    #         print(f"{application_name} closed successfully")
    #     except Exception as e:
    #         print(f"Error while closing {application_name}: {e}")
    # else:
    #     print(
    #         f"Sorry, I couldn't find {application_name} in the list of supported applications")


# def email():
#     try:
#         speak("What should I say?")
#         content = takecommand()
#         speak("whome should i send")
#         to = input()
#         sendEmail(to, content)
#         speak("Email has been sent !")
#     except Exception as e:
#         print(e)
#         speak("I am not able to send this email")


# def sendEmail(to, content):
#     server = smtplib.SMTP('smtp.gmail.com', 587)
#     server.ehlo()
#     server.starttls()
#     server.login('surajgupta3101@gmail.com', '@guptasuraj31')
#     server.sendmail('surajgupta3101@gmail.com', to, content)
#     server.close()
#     pass

def sendAnEmail():
    now = datetime.datetime.now()
    SERVER = "smtp.gmail.com"
    PORT = 587
    FROM = "surajgupta3101@gmail.com"
    TO = "surajgupta3101@gmail.com"
    PASS = '@guptasuraj31'

    msg = MIMEMultipart()
    msg['Subject'] = "Email From Suraj Gupta [Automated Email]" + \
        ' ' + str(now.day) + '-' + str(now.month) + '-' + str(now.year)
    msg['From'] = FROM
    msg['To'] = TO
    speak("What should i say?")
    print("Content: ")
    content = takecommand().lower()

    msg.attach(MIMEText(content.lower(), 'html'))

    print("Initializing server")

    server = smtplib.SMTP(SERVER, PORT)
    server.set_debuglevel(1)
    server.ehlo()
    server.starttls()
    server.login(FROM, PASS)
    server.sendmail(FROM, TO, msg.as_string())

    speak('Email Sent')
    server.quit()


# def findspeed():
#     speak("Checking Speed...")
#     speed = speedtest.Speedtest()
#     speed.get_best_server()
#     downloading = speed.download()
#     correctDown = int(downloading/800000)
#     uploading = speed.upload()
#     correctUpload = int(uploading/800000)

#     if 'uploading' in query:
#         speak(f"The uploading speed is: {correctUpload} mbps")

#     elif 'downloading' in query:
#         speak(f"The downloading speed is: {correctDown} mbps")

#     else:
#         speak(
#             f"The Downloading speed is {correctDown} and The uploading speed is {correctUpload}")


#         elif "send message on whatsapp" in query:
#             # Format: (number, msg which you want to send, time stamp(hr, minute))
#             # kit.sendwhatmsg("+917217571900", "this is testing protocol", 00, 17)

#             # Method-2: Send what you want to say
#             # speak("To whom you want to send the message")
#             # whom = takecommand().lower()
#             speak("What message you want to send")
#             message = takecommand().lower()
#             # num = takecommand().lower()
#             # kit.sendwhatmsg("+91",f"{whom}", f"{message}", 16,10)
#             # kit.sendwhatmsg("+919065090606", f"{message}", 16,16)
#             # kit.sendwhatmsg("+917217571900", "hey sushant")

#             current_time = datetime.datetime.now()
#             hr = current_time.hour
#             print(hr)
#             mnt = current_time.minute
#             print(mnt)
#             # kit.sendwhatmsg("+917217571900", f"{message}", hr, mnt+)
#             # Error: country code is missing in phone number

#         elif "play song on youtube" in query:
#             speak("Which song you want to play")
#             song = takecommand().lower()
#             pywhatkit.playonyt(f"{song}")

#             # Methd-2:
#             # kit.playonyt(song name)


class AssistanceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Voice Assistant")
        self.root.geometry('600x600')

        self.bg = ImageTk.PhotoImage(file="images\\background.png")
        bg = Label(self.root, image=self.bg).place(x=0, y=0)

        self.center = ImageTk.PhotoImage(file="images\\frame_image.jpg")
        left = Label(self.root, image=self.center).place(
            x=100, y=100, width=400, height=400)

        start = Button(self.root, text='START', font=("times new roman", 14), command=self.start_option).place(x=150,
                                                                                                               y=520)
        close = Button(self.root, text='CLOSE', font=("times new roman", 14), command=self.close_window).place(x=350,
                                                                                                               y=520)

        self.assistant_active = False
        self.assistant_thread = None

    def start_option(self):
        self.assistant_active = True
        self.assistant_thread = threading.Thread(target=self.run_assistant)
        self.assistant_thread.start()

    def run_assistant(self):
        listener = sr.Recognizer()
        engine = pyttsx3.init()

        def speak(text):
            engine.say(text)
            engine.runAndWait()

        def start():
            hour = datetime.datetime.now().hour
            if 0 <= hour < 12:
                wish = "Good Morning!"
            elif 12 <= hour < 18:
                wish = "Good Afternoon!"
            else:
                wish = "Good Evening!"
            speak(
                f'Hello Sir, {wish} I am your voice assistant. Please tell me how may I help you')

        def take_command():
            try:
                with sr.Microphone() as data_taker:
                    print("Say Something")
                    voice = listener.listen(data_taker)
                    instruction = listener.recognize_google(voice)
                    instruction = instruction.lower()
                    return instruction
            except sr.UnknownValueError:
                pass

        def run_command():
            def clear(): return os.system('cls')

            clear()
            # wish()
            # username()
            while self.assistant_active:
                query = take_command()
                print(query)
                try:

                    if 'open' in query:
                        query = query.replace("open", "")
                        open_application(query.lower())

                    elif "close" in query:
                        query = query.replace("close", "")
                        close_application(query.lower())

                    # f2-> done
                    elif "ip address" in query:
                        find_ip()

                    # f5 -> done
                    elif "find the temperature" in query or 'find out the temperature' in query or 'tell me the temperature' in query:
                        find_temp()

                    #  f6 -> done
                    elif 'according to google' in query:
                        findOnGoogle(query)

                    # f7 -> done
                    elif 'set the time to wake up' in query or 'set wake up time' in query:
                        timeToWakeUp()

                    # f8 -> done
                    elif "wikipedia" in query:
                        resultOnWikipedia(query)

                    # f9 -> done
                    elif 'current time' in query:
                        findCurrentTime()

                    # f10 -> done
                    elif "change my name to" in query or "change name" in query or 'change my name' in query:
                        changeName(query)

                    # f11 -> done
                    elif "joke" in query:
                        tellJoke()

                    # f13 -> done
                    elif 'lock the window' in query:
                        lockWindow()

                        # f14
                    elif 'shutdown the system' in query:
                        shutDown()

                    # f15 -> done
                    elif 'empty the recycle bin' in query or 'empty recycle bin' in query or 'empty bin' in query:
                        emptyBin()

                    # f16 -> done
                    elif "don't listen" in query or 'stop listening' in query or "hold for some time" in query or "hold" in query or 'stop listen' in query:
                        StopListen()

                    # f17 -> done
                    elif 'location of' in query or 'locate' in query or 'find the location of' in query:
                        locationOnMap(query)

                    # f18 - done
                    elif 'click a photo' in query or 'take a photo' in query or 'take photo' in query or 'take selfie' in query:
                        capturePhoto()

                        # f19
                    elif 'restart the system' in query or 'restart system' in query:
                        restartSystem()

                        # f20
                    elif 'write something' in query or 'write a note' in query or 'write note' in query or 'write information' in query or 'note something' in query:
                        writeNote()

                        # f21
                    elif 'show note' in query:
                        showNote()

                        # f22 -> done
                    elif 'find the movie' in query or 'find movie' in query or 'find the detail of movie' in query or 'movies' in query:
                        findMovies()

                        # f23
                    elif 'send mail' in query or 'send an email' in query or 'send email' in query:
                        sendAnEmail()

                        # f24 -> done
                    elif 'image' in query:
                        detect_image()

                    elif 'current time' in query:
                        time = datetime.datetime.now().strftime('%I: %M')
                        speak('current time is' + time)

                    elif 'shutdown' in query or 'session terminate' in query:
                        speak('Terminating the session')
                        self.assistant_active = False
                        self.close_window()

                    else:
                        speak('I did not understand, can you repeat again')
                except:
                    speak('Waiting for your response')

        start()
        run_command()

    def close_window(self):
        self.assistant_active = False
        if self.assistant_thread and self.assistant_thread.is_alive():
            self.assistant_thread.join()
        self.root.quit()


root = Tk()
obj = AssistanceGUI(root)
root.mainloop()
