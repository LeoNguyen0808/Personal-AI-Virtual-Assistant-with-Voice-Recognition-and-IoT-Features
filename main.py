#process, access the file system
from ast import Str
import ctypes
from fileinput import filename
import os
import pickle
from telnetlib import EC
from turtle import clear
from fastapi import Request
#convert text to speech
from gtts import gTTS
#open sound
import playsound
#detect user voice
import speech_recognition
#take date time data
from time import strftime
import time
import datetime
#random selection
import random
#for website
import re
import webbrowser
#command to take web information
import requests
import json
import wikipedia
#support for searching website,open youtube
from webdriver_manager.chrome import ChromeDriverManager
from youtubesearchpython import SearchVideos
import urllib
import urllib.request as urllib
#tkinter modules
from tkinter import Tk, RIGHT, BOTH, RAISED
from tkinter.ttk import Frame, Button, Style
from tkinter import *
from PIL import Image, ImageTk, ImageSequence
import tkinter.messagebox as mbox
import tkinter as tk
# Google API vs Gmail and Google Calendar
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from calendar import calendar
import datetime
import pickle
import os.path

#---------------------------------PREPARING STEP---------------------------------------------#
path=ChromeDriverManager().install()
root = tk.Tk()
# Change application icon
root.tk.call('wm', 'iconphoto', root._w, tk.PhotoImage(file='FYP - AI Virtual Assistant//345.png')) 
text_area = Text(root, height=25, width=45,wrap = WORD)
scroll = Scrollbar(root, command=text_area.yview)
language = 'en'

# LEO will speak with the support of gTTS library
def speak(text):
    print("Leo:  {}".format(text))
    text_area.insert(INSERT,"Leo: "+text+"\n")
    tts = gTTS(text=text, lang= language, slow=False)
    tts.save("FYP - AI Virtual Assistant//sound.mp3")
    playsound.playsound("FYP - AI Virtual Assistant//sound.mp3", False)
    os.remove("FYP - AI Virtual Assistant//sound.mp3")

# The system will recognize the user's voice and convert it to text format.
def get_audio():
    playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
    time.sleep(3)
    r = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as source:
        print("You: ")
        audio = r.listen(source, phrase_time_limit=6)
        try:
            text = r.recognize_google(audio, language="en-US")
            print(text)
            return text.lower()
        except:
            print("\n")
            return ""
#-------------------------------------------END-----------------------------------------#


#-----------------------------------------GREETING FUNCTIONS----------------------------#       
def hello():
    day_time = int(strftime('%H'))
    if day_time < 12:
        img = Image.open("FYP - AI Virtual Assistant//morning.jpg")
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        speak("Good Morning! How can I help you?")
    elif 12 <= day_time < 18:
        img = Image.open("FYP - AI Virtual Assistant//afternoon.jpg")
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        speak("Good Afternoon! How can I help you?")
    else:
        img = Image.open("FYP - AI Virtual Assistant//evening.jpg")
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        speak("Good Evening! How can I help you?")
    root.update()
    time.sleep(5)
#---------------------------------------------------END---------------------------------#


#----------------------------------------DATETIME FUNCTIONS-----------------------------#
def get_time(text):
    img= (Image.open("FYP - AI Virtual Assistant//datetime.jpg"))
    #Resize the Image using resize method
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    now = datetime.datetime.now()
    now1 = datetime.datetime.now().strftime("%A")
    if "time" in text:
        speak('Now is %d:%d:%d' % (now.hour, now.minute, now.second))
    elif "date" or "day" in text:
        speak("Today is\t" + now1 + "\t%d/%d/%d" % (now.day, now.month, now.year))
    else:
        speak("I can't understand what your means?\n" 
              "Can you speak again?")
        time.sleep(3)
    root.update()
    time.sleep(5)
#----------------------------------------END----------------------------------------------#


#------------------------------------OPEN APPS/APPLICATIONS-------------------------------#
def open_application():
    global img
    img= (Image.open("FYP - AI Virtual Assistant//c.gif"))
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
          
    speak('Which app/application should I open for you?')
    root.update()
    time.sleep(4)
    text = get_audio()
    text_area.insert(INSERT,"You: "+text+"\n")
    root.update()
    if "access" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk"')
        speak("Microsoft Access opened")
    elif "adobe" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Creative Cloud.lnk"')
        speak("Adobe Creative Cloud opened")
    elif "excel" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk"')
        speak("Microsoft Excel opened")
    elif "chrome" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"')
        speak("Google Chrome opened")
    elif "edge" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk"')
        speak("Microsoft Edge opened")
    elif "onenote" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"')
        speak("Microsoft OneNote opened")
    elif "onedrive" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk"')
        speak("Microsoft OneDrive opened")
    elif "outlook" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk"')
        speak("Microsoft Outlook opened")
    elif "powerpoint" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk"')
        speak("Microsoft PowerPoint opened")
    elif "tableau" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Tableau 2021.4.lnk"')
        speak("Tableau opened")
    elif "word" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"')
        speak("Microsoft Word opened")
    elif "packet tracer" in text:
        os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Cisco Packet Tracer\Cisco Packet Tracer.lnk"')
        speak("Cisco Packet Tracer opened")
    elif "control panel" in text:
        os.startfile('"C:\\Users\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\System Tools\\Control Panel.lnk"')
        speak("Control Panel opened")
    elif "command prompt" in text:
        os.startfile('"C:\\Users\\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\System Tools\\Command Prompt.lnk"')
        speak("Command Prompt opened")
    elif "visual studio code" in text:
        os.startfile('"C:\\Users\\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Visual Studio Code\Visual Studio Code.lnk"')
        speak("Visual Studio Code opened")
    elif "grammarly" in text:
        os.startfile('"C:\\Users\\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Grammarly\\Grammarly.lnk"')
        speak("Grammarly opened")
    elif "discord" in text:
        os.startfile('"C:\\Users\\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Discord Inc\\Discord.lnk"')
        speak("Discord opened")
    elif "orange" in text:
        os.startfile('"C:\\Users\\Dell\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Orange\\Orange.lnk"')
        speak("Orange opened")
    elif "close" in text:
        ai_brain="My pleasure when I can help you.\n See you again. Have a good day!"
        speak(ai_brain)
        root.update()
        time.sleep(4)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(0.5)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(0.5)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(1)
        exit()
    elif "exit" in text:
        ai_brain="My pleasure when I can help you.\n See you again. Have a good day!"
        speak(ai_brain)
        root.update()
        time.sleep(4)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(0.5)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(0.5)
        playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
        time.sleep(1)
        exit()
    else:
        speak(
            "Please check it is correct app name.\n "
            "Or Your requested app/application is not install.\n"
            "Please try to open another app or install it!")
        root.update()
        time.sleep(12)
        open_application()
    text_area.insert(INSERT,"_____________________________________________")
#------------------------------------------------END-----------------------------------------#    


#-------------------------------------------OPEN WEBSITE-------------------------------------#
def open_website():
    img = Image.open("FYP - AI Virtual Assistant//website.jpg")  
    
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    speak('Which website should I open for you?\n' 
          'Please following command "Open" + the domain format with ".com",".org" and so on.')
    root.update()
    time.sleep(12)
    text = get_audio()
    text_area.insert(INSERT,"You: "+text+"\n")
    root.update()
    reg_ex = re.search('open (.+)', text)
    
    if reg_ex:
        domain = reg_ex.group(1)
        url = 'https://www.' + domain
        webbrowser.open(url)
        speak("Your requested website opened.")
    else:
        webbrowser.open("https://www.google.com.sg/")
        speak("I don't get your requested website.\n" 
              "But I will open Chrome for you to searching.")
        root.update()
        time.sleep(5)
    text_area.insert(INSERT,"_____________________________________________")
#----------------------------------------------END------------------------------------------------#


#-------------------------------------------GOOGLE SEARCHING--------------------------------------#
def open_google_and_search():
    img = Image.open("FYP - AI Virtual Assistant//google.jpg")   
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    speak('What should I search for you?')
    root.update()
    time.sleep(2.5)
    text = get_audio()
    text_area.insert(INSERT,"You: "+text+"\n")
    root.update()
    
    if text == "":
        speak("Error. I can't get your request.\n" 
            "Please speak again.")
        root.update()
        time.sleep(5)
        
        speak('What should I search for you?')
        root.update()
        time.sleep(2.5)
        text1 = get_audio()
        text_area.insert(INSERT,"You: "+text1+"\n")
        root.update()
        if text1 == "":
        
            speak("Error. I can't get your request.\n" 
                "Please try again with Google search or other functions.")
            root.update()
            time.sleep(7)
            
        else:
            search= text1.lower()
            url= (f"http://www.google.com/search?q={search}")
            webbrowser.get().open(url)
            speak(f'Here is your {search} on Google.')
            root.update()    
    else:
        search= text.lower()
        url= (f"http://www.google.com/search?q={search}")
        webbrowser.get().open(url)
        speak(f'Here is your {search} on Google.')
        root.update()
    text_area.insert(INSERT,"_____________________________________________")    
#---------------------------------------------------END----------------------------------------------#


#----------------------------------------YOUTUBE OPENING/SEARCHING-----------------------------------#
def youtube_search():
    img = Image.open("FYP - AI Virtual Assistant//youtube.jpg")    
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)

    speak('Please let me know what you want to search/open in YouTube')
    root.update()
    time.sleep(4)
    text = get_audio()
    text_area.insert(INSERT,"You: "+text+"\n")
    root.update()
    
    if text == "":
        speak("Error. You don't say anything, please speak again.")
        time.sleep(6)
        root.update()
        
        speak('Please let me know what you want to search/open in YouTube')
        root.update()
        time.sleep(5)
        text = get_audio()
        text_area.insert(INSERT,"You: "+text+"\n")
        root.update()
        if text == "":
            speak("Error. You don't say anything.\n "
                  "Please try again with YouTube opening or other functions.")
            root.update()
            time.sleep(6)
        else:
            try:
                search = SearchVideos(text, offset = 1, mode = "json", max_results = 20).result()
                data = json.loads(search)
                url = data["search_result"][0]["link"]
                print(url)
                webbrowser.open(url)
                if "album" or "song" in text:
                    speak("Your requested album/sing opened.")
                        
                elif "film" in text:
                    speak("Your requested film opened.")
                else:
                    speak("Your requestment completed.")
            except:
                speak("Error. Your request is not available.\n "
                    "Please try again with YouTube opening or other functions.")
    else:
        try:
            search = SearchVideos(text, offset = 1, mode = "json", max_results = 30).result()
            data = json.loads(search)
            url = data["search_result"][0]["link"]
            print(url)
            webbrowser.open(url)
            if "album" or "song" in text:
                speak("Your requested album/sing opened.")
                    
            elif "film" in text:
                speak("Your requested film opened.")
            else:
                speak("Your requestment completed.")
        except:
            speak("Error. Your request is not available.\n "
                  "Please try again with YouTube opening or other functions.")
    text_area.insert(INSERT,"_____________________________________________")    
#-----------------------------------------------END--------------------------------------------#        

 
#---------------------------------------CHANGE COMPUTER/LAPTOP BACKGROUND----------------------#           
def change_wallpaper():
    img = Image.open("FYP - AI Virtual Assistant//a.gif") 
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    api_key = 'HP4j8qO5NOcO_V7S7V7fdB2rQufN52fFi5qKalcjA9w'
    url = 'https://api.unsplash.com/photos/random?client_id=' + \
        api_key  # pic from unspalsh.com
    f = urllib.urlopen(url)
    json_string = f.read()
    f.close()
    parsed_json = json.loads(json_string)
    photo = parsed_json['urls']['full']
    # Location where we download the image to.
    urllib.urlretrieve(photo, "C:/Users/Dell/Downloads/a.png")
    image=os.path.join("C:/Users/Dell/Downloads/a.png")
    ctypes.windll.user32.SystemParametersInfoW(20,0,image,3)
    speak('Your computer/laptop background changed.')
    text_area.insert(INSERT,"_____________________________________________")
#--------------------------------------------------END---------------------------------------#


#-------------------------------------- DISPLAY CITY'S WEATHER-------------------------------#
def current_weather():
    img = Image.open("FYP - AI Virtual Assistant//weather.jpeg") 
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    speak("Where you want to see the weather?")
    root.update()
    time.sleep(3)
    ow_url = "http://api.openweathermap.org/data/2.5/weather?"
    city = get_audio()
    text_area.insert(INSERT,"You: "+city+"\n")
    if city=="":
        speak("Sorry. I can't get your location. Please speak again!")
        root.update()
        time.sleep(5)
        current_weather()
    else:
        api_key = "fe8d8c65cf345889139d8e545f57819a"
        call_url = ow_url + "appid=" + api_key + "&q=" + city + "&units=metric"
        response = requests.get(call_url)
        data = response.json()
        if data["cod"] != "404":
            city_res = data["main"]
            current_humidity = city_res["humidity"]
            current_temperature = city_res["temp"]
            current_pressure = city_res["pressure"]
            suntime = data["sys"]
            sunrise = datetime.datetime.fromtimestamp(suntime["sunrise"])
            sunset = datetime.datetime.fromtimestamp(suntime["sunset"])

            now = datetime.datetime.now()
            current_time = now.strftime("%H:%M:%S")
            content = """Now is {time}.\nToday is {day}/{month}/{year}.
Sunrise happens {hourrise} hours {minrise} minutes.
Sunset happens {hourset} hours {minset} minutes.
The average temperature at  {temp} °C.
Humidity is {humidity}%.
Air pressure is {pressure} Hz.""".format(time= current_time, day = now.day,month = now.month, year= now.year,
                                        hourrise = sunrise.hour, minrise = sunrise.minute,
                                        hourset = sunset.hour, minset = sunset.minute, 
                                        temp = current_temperature, humidity = current_humidity,
                                        pressure =current_pressure)
            speak(content)
            root.update()
            time.sleep(27)
        else:
            speak("I can't get your location. Can you speak again?")
            root.update()
            time.sleep(5)
            current_weather()
    text_area.insert(INSERT,"_____________________________________________")
#---------------------------------------------------------END-----------------------------------------------#


#----------------------------------------WIKIPEDIA DEFINITION-----------------------------#
def get_definition():
    try:
        img = Image.open("FYP - AI Virtual Assistant//wikipedia.jpg") 
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        speak("Welcome to Wikipedia definition service.\nWhat you want to get definition?")
        root.update()
        time.sleep(6)
        text = get_audio()
        text_area.insert(INSERT,"You: "+text+"\n")
        root.update()
        contents = wikipedia.summary(text).split('\n')
        speak(contents[0])
        root.update()
        time.sleep(50)
        for content in contents[1:]:
            speak("Do you want to get more new definition?")
            ans = get_audio()
            text_area.insert(INSERT,"You: "+ans+"\n")
            root.update()
            if "yes" not in ans:
                break    
            speak(content)
            time.sleep(50)

        speak('Thank you for listening.')
        time.sleep(2)
    except:
        speak("I can't get what you want to get definition. Please try again.")
        time.sleep(7)
    text_area.insert(INSERT,"_____________________________________________")
#--------------------------------------------END-------------------------------------------------------#    


#-------------------------------------GOOGLE CALENDAR WITH UPCOMING EVENTS-----------------------------#
SCOPES1 = ['https://www.googleapis.com/auth/calendar.readonly']
def authenticate_google():
    
    creds = None
    if os.path.exists('token1.pickle'):
        with open('token1.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'FYP - AI Virtual Assistant//calendargg.json', SCOPES1)
            creds = flow.run_local_server(port=0)

        with open('token1.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def get_events(n, service):
    img = Image.open("FYP - AI Virtual Assistant//events.jpg") 
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    speak ("Welcome to Google Calendar events!")
    root.update()
    time.sleep(4)
    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat()+ 'Z' # 'Z' indicates UTC time
    events = 'Getting your upcoming events...'
    speak(events)
    root.update()
    time.sleep(4)
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=n, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
        root.update()
        time.sleep(4)
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        summary = event.get("summary", "No summary found!")
        text = start +"\t" + summary
        speak(text)
        root.update()
        time.sleep (10)
    speak('These are all the upcoming events you requested.')
    root.update()    
    text_area.insert(INSERT,"_____________________________________________")   

#------------------------------ EMAIL FUNCTIONS --------------------------------------------#
#Check authenticate gmail account if the users haven't access LEO get the email infor.
def authenticate_gmail():
        SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
        """Shows basic usage of the Gmail API.
        Lists the user's Gmail labels.
        """
        creds = None
        
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
                
        # If there are no (valid) credentials available,
        # Let the user log in to the google account with Gmail apps
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'FYP - AI Virtual Assistant//email.json', SCOPES)
                creds = flow.run_local_server(port = 0)
                
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('gmail', 'v1', credentials=creds)
        return service
# Get the 100 nearest emails.  
def check_mails(service):
        img = Image.open("FYP - AI Virtual Assistant//new_email.jpg")   
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        
        speak("Welcome to newest mail service!")
        root.update()
        time.sleep(3)
        
        # Call the Gmail API
        results = service.users().messages().list(userId = 'me').execute()
        # The above code will get emails from primary
        # inbox which are unread
        messages = results.get('messages')

        if not messages:
            # if no new emails
            speak('No messages found.')
            root.update()
        else:
            m=""
            
            # if email found
            speak("{} newest emails found".format(len(messages)))
            root.update()
            time.sleep(3)
            
            speak("If you want to read any particular email just say 'read' or 'okay'."
                  " Just say anything to pass/skip email and come to the next email.")
            root.update()
            time.sleep(12)
        
            for message in messages:
                msg=service.users().messages().get(userId='me',
                id = message['id'],format="full").execute()

                for add in msg['payload']['headers']:
                    if add['name']=="From":
                        # fetching sender's email name
                        a=str(add['value'].split("<")[0])
                        print(a)

                        speak("Email from\t" + a)
                        root.update()
                        time.sleep(2)
                        text = get_audio()
                        
                        text_area.insert(INSERT,"You: "+text+"\n")
                        root.update()
                        if text == "okay" or text == "read":
                            speak(msg.get('snippet'))
                            root.update()
                            time.sleep(15)
                        elif text =="stop":
                            speak("I'm stopping...")
                            root.update()
                            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
                            time.sleep(0.5)
                            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
                            time.sleep(0.5)
                            text_area.insert(INSERT,"_____________________________________________")
                            break
                        else:
                            speak("Email passed")
                            root.update()
                            time.sleep(2)
                         			

# Get all the unbox email begin when user's allow Leo to access their emails data.
def unbox_mails(service):
        img = Image.open("FYP - AI Virtual Assistant//unread.jpg")   
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        
        speak("Welcome to unread mail service!")
        root.update()
        time.sleep(3)
        # fetching emails of today's date
        today = (datetime.date.today())

        today_main = today.strftime('%Y/%m/%d')

        # Call the Gmail API
        results = service.users().messages().list(userId = 'me',
                                                  labelIds=["INBOX","UNREAD"],
											q="after:{0} and category:Primary".format(today_main)).execute()
        # The above code will get emails from primary
        # inbox which are unread
        messages = results.get('messages', [])

        if not messages:
            # if no new emails
            speak('No messages found.')
            root.update()
        else:
            # if email found
            speak("{} unread emails found".format(len(messages)))
            root.update()
            time.sleep(3)
            
            speak("If you want to read any particular email just say 'read' or 'okay'."
                  " Just say anything to come next email.")
            root.update()
            time.sleep(12)
        
        for message in messages:
            msg=service.users().messages().get(userId='me',
            id = message['id'],format="full").execute()

            for add in msg['payload']['headers']:
                if add['name']=="From":
                                # fetching sender's email name
                    a=str(add['value'].split("<")[0])
                    print(a)

                    speak("Email from\t" + a)
                    root.update()
                    time.sleep(2)
                    text = get_audio()
                    root.update()
                    text_area.insert(INSERT,"You: "+text+"\n")
                    root.update()
                    
                    if text == "okay" or text == "read":
                        speak(msg.get('snippet'))
                        root.update()
                        time.sleep(25)
                    
                    else:
                        speak("Email passed")
                        root.update()
                        time.sleep(2)
#---------------------------------- END --------------------------------------------------#


#---------------------------------- IOT CONTROL ------------------------------------------#
def IoT_control():
    img = Image.open("FYP - AI Virtual Assistant//Light-control.jpg") 
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)

    speak('Currently I can able to control the LED and Fan devices. So, What you want to do?')
    root.update()
    time.sleep(8)
    text = get_audio()
    text_area.insert(INSERT,"You: "+text+"\n")
    root.update()
    if "on fan" in text or "on the fan" in text:
        webbrowser.open('http://192.168.0.143/FAN=ON',new=1) 
        speak("The fan is on for you")
        root.update()
    elif "off fan" in text or "off the fan" in text:
        webbrowser.open('http://192.168.0.143/FAN=OFF',new=1) 
        speak("The fan is off for you")
        root.update()
    elif "on light" in text or "on the light" in text:
        webbrowser.open('http://192.168.0.143/LED=ON',new=1)
        speak("The light is on for you")
        root.update()
    elif "off light" in text or "off the light" in text:
        webbrowser.open('http://192.168.0.143/LED=OFF',new=1) 
        speak("The light is off for you")
        root.update()
    else:
        speak("Sorry I can't get what you mean")
        return 
    text_area.insert(INSERT,"_____________________________________________")
#--------------------------------------- END IOT FUNCTIONS ---------------------------------#
  
        
# DISPLAY THE LEO FUNCTIONS         
def func():
    img = Image.open("FYP - AI Virtual Assistant//function.gif")   
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    
    content="""
    My functions include:
    1. Greetings
    2. Show Date & Time 
    3. Open apps/applications
    4. Open the Intenet website 
    5. Search information from Google
    6. Open musics or films in Youtube 
    7. Display the weather details
    8. Change the computer wallpaper
    9. Email Service
    10. Upcoming Events from Google Calendar
    11. Wikipedia Definition
    12. IoT control with Fan and LED
    """
    speak(content)
    root.update()
    time.sleep(45)

# Display the IoT image
def iot_image():
    img = Image.open("FYP - AI Virtual Assistant//Light-control.jpg") 
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
       
# Contact Us Icon Clickable Link
def fb_click(event=None):
    webbrowser.open("https://www.facebook.com/messages/new")

def insta_click(event=None):
    webbrowser.open("https://www.instagram.com/_minh_nhat_nguyen_/")

def wapp_click(event=None):
    webbrowser.open("https://web.whatsapp.com/") 

# Change the Background Communication Area    
def color():
    mylist = ["#009999","black","green","grey","blue","orange","#cc0099","#00ff00","brown"]
    aa=random.choice(mylist)
    text_area.tag_add("here", "1.0", "100000.0")
    text_area.tag_config("here", background = aa)
    root.update()

# Change the Font Communication Area
def color1():
    mylist1 = ["yellow","#0000ff","white","#00ff00","black"]
    bb=random.choice(mylist1)
    text_area.tag_add("here", "1.0", "100000.0")
    text_area.tag_config("here",foreground = bb)
    root.update()

# LEO Functions display
def func_info():
    mbox.showinfo("Leo's Functions", "1. Greetings\n"
    "2. Show Date & Time\n" 
    "3. Open apps/applications\n"
    "4. Open the Intenet website\n" 
    "5. Search information from Google\n"
    "6. Open musics or films in Youtube\n" 
    "7. Display the weather details\n"
    "8. Change the computer wallpaper\n"
    "9. Email Service\n"
    "10. Upcoming Events from Google Calendar\n"
    "11. Wikipedia Definition\n"
    "12. IoT control with Fan and LED\n")

# 'About Us' Content
def info():
    mbox.showinfo("Introduce", "- Press 'Micro' button to start communication with AI.\n"
                  "- Click 'Refresh' to start new conversation.\n"
                  "- When the 'Ping' sound appear that means Leo is listening...\n"
                  "- When the cross line appear that mean Leo is stopping and click 'Micro' button to continue.\n"
                  "- Speak 'Stop/Pause' to temporary stop the conversation.\n"
                  "- Click or speak 'Exit' button to close Leo application.")
    
# Exit the application    
def r_set():
    img = Image.open("FYP - AI Virtual Assistant//mainbackground.jpg")    
    resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
    image_1= ImageTk.PhotoImage(resized_image)
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=65)
    text_area.delete("1.0", "1000000000.0")

#----------------------------------- COMMUNICATION GUIDELINES ---------------------------#
def ham_main():
    i = 0
    you=""
    ai_brain=""
    while i<3:
        r = speech_recognition.Recognizer()
        with speech_recognition.Microphone() as source:
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(3)
            print("Leo:  Listening ...")
            audio = r.listen(source, phrase_time_limit=3)
            # print("Leo:  ...")
        try:
            you = r.recognize_google(audio, language="en-US")
            print("\nYou: "+ you)	
            you = you.lower()
        except:
            i +=1
            

        text_area.insert(INSERT,"You: "+you+"\n")
        root.update()

        if "hi" in you or "hello" in you or "halo" in you or'good moring' in you or 'good afternoon' in you or'good evening' in you or'good night in you' in you:
            hello()
        elif "weather" in you or "temperature" in you:
            current_weather()
            break
        elif "time" in you or "date" in you or "day" in you:
            get_time(you)
        elif "open application" in you or "open app" in you or "open the application" in you or "open the app" in you:
            open_application()
            break
        elif "open website" in you:
            open_website()
            break
        elif "google" in you:
            open_google_and_search()
            break
        elif " music" in you or "film" in you or "youtube" in you or "sing" in you:
            youtube_search()
            break
        elif "function" in you or "feature" in you:
            func()
            speak("Which function/feature can I do for you?")
            root.update()
            time.sleep(4)
        elif "background color" in you:
            color()
        elif "font color" in you:
            color1()
        elif "background" in you or "wallpaper" in you or "background desktop" in you or "deskop wallpaper"in you or"desktop background" in you:
            change_wallpaper()
            break
        elif "events" in you or "calendar" in you or "upcoming" in you:
            authenticate_google()
            speak('Please let me know how many upcoming events you want to see?')
            root.update()
            time.sleep(4)
            text = get_audio()
            text_area.insert(INSERT,"You: "+ text +"\n")
            n=str(text)
            root.update()
            if n.isdigit():
                get_events(n,authenticate_google())
                break
            else:
                speak("Please speak command correctly with the number format.Thank you.")
                text_area.insert(INSERT,"_____________________________________________")
                break
        elif "iot control" in you or 'iot' in you or 'control' in you:
            IoT_control()
            break
        elif "on fan" in you or "on the fan" in you or "fan on" in you:
            iot_image()
            webbrowser.open('http://192.168.0.143/FAN=ON',new=1) 
            speak("The fan is on for you.")
            text_area.insert(INSERT,"_____________________________________________")
            root.update()
            break
        elif "off fan" in you or "off the fan" in you or "fan off" in you:
            iot_image()
            webbrowser.open('http://192.168.0.143/FAN=OFF',new=1) 
            speak("The fan is off for you.")
            text_area.insert(INSERT,"_____________________________________________")
            root.update()
            break
        elif "on light" in you or "on the light" in you or "light on" in you:
            iot_image()
            webbrowser.open('http://192.168.0.143/LED=ON',new=1)
            speak("The light is on for you.")
            text_area.insert(INSERT,"_____________________________________________")
            root.update()
            break
        elif "off light" in you or "off the light" in you or "light off" in you:
            iot_image()
            webbrowser.open('http://192.168.0.143/LED=OFF',new=1) 
            speak("The light is off for you.")
            text_area.insert(INSERT,"_____________________________________________")
            root.update()
            break
        elif "definition" in you or "wikipedia" in you:
            get_definition()
            speak("Which function/feature can I do for you?")
            root.update()
            time.sleep(4)
        elif "email" in you or "gmail" in you:
            speak("Do you want to check newest or unread email?")
            root.update()
            time.sleep(4)
            text = get_audio()
            text_area.insert(INSERT,"You: "+ text +"\n")
            root.update()
            if "new" in text:
                authenticate_gmail()
                check_mails(authenticate_gmail())
                text_area.insert(INSERT,"_____________________________________________")
                break
            elif "unread" in text or "unbox" in text:
                authenticate_gmail()
                unbox_mails(authenticate_gmail())
                text_area.insert(INSERT,"_____________________________________________")
                break
            else: 
                speak("Sorry, I can't get your word.\n Please speak new or unread to step in email service.")
                text_area.insert(INSERT,"_____________________________________________")
                break
        elif "stop" in you or "pause" in you:
            speak("I'm stopping...")
            root.update()
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(0.5)
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(0.5)
            text_area.insert(INSERT,"_____________________________________________")
            break
        elif "exit" in you or "close" in you:
            ai_brain="My pleasure when I can help you.\n See you again. Have a good day!"
            speak(ai_brain)
            root.update()
            time.sleep(4)
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(0.5)
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(0.5)
            playsound.playsound("FYP - AI Virtual Assistant//Ping.mp3", False)
            time.sleep(1)
            exit()
        elif " " in you:
            ai_brain = ("I can't get your word.\n" 
                       "Can you speak again?")
            speak(ai_brain)
            root.update()
            time.sleep(4)
        else:
            ai_brain = ("I can't understand what you means\n" 
                       "Can you speak again?")
            speak(ai_brain)
            root.update()
            time.sleep(4)

        text_area.insert(INSERT,"_____________________________________________")
        you=""
#--------------------------------------------- END -----------------------------------------------------#


#-------------------------------------- SET UP THE APPLICATION LAYOUT-----------------------------------#
class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent) 
        self.parent = parent
        self.initUI()
    
    def initUI(self):
        self.parent.title("LEO - Personal AI Virtual Assistant with Voice Recognition and IoT Features")
        self.style = Style()

        scroll.pack(side=RIGHT, fill=Y)
        text_area.configure(yscrollcommand=scroll.set)
        text_area.pack(side=RIGHT)
        frame = Frame(self, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        self.pack(fill=BOTH, expand=True)
        
        # MENU BAR
        menubar = Menu(root)
        filemenu = Menu(menubar, tearoff=False, background="#A9D09E")
        filemenu.add_command(label="Refresh", command=r_set)
        filemenu.add_command(label="Undo", command=ham_main)
        filemenu.add_command(label="Redo", command=color)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="Edit", menu=filemenu)
        
        settingmenu = Menu(menubar, tearoff=False, background="#A9D09E" )
        settingmenu.add_command(label="Change background color", command=color)
        settingmenu.add_separator()
        settingmenu.add_command(label="Change text color", command=color1)
        menubar.add_cascade(label="Settings", menu=settingmenu)
        
        helpmenu = Menu(menubar, tearoff=False, background="#A9D09E" )
        helpmenu.add_command(label="Help Index", command=info)
        helpmenu.add_command(label="About...", command=info)
        helpmenu.add_separator()
        helpmenu.add_command(label="Leo's Functions", command=func_info)
        menubar.add_cascade(label="Help", menu=helpmenu)
        root.config(menu=menubar)
        
        # MAIN BACKGROUND DISPLAY 
        img= (Image.open("FYP - AI Virtual Assistant//mainbackground.jpg"))
        resized_image= img.resize((600,400), Image.Resampling.LANCZOS)
        image_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(self, image=image_1)
        label1.image = image_1
        label1.place(x=7, y=65)
        
        # PROFILE BUTTON 
        nmn= (Image.open("FYP - AI Virtual Assistant//profile.jpg"))
        resized_image= nmn.resize((35,35), Image.Resampling.LANCZOS)
        nmn_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(self, image=nmn_1)
        label1.image = nmn_1
        b = tk.Button(self, image=nmn_1, command=insta_click)
        b.place(x=15, y=8)
        
        # SETTING THE TEXT CONTENT
        l = Label(root, text='꧁ Communication Area ꧂', fg='black',bg="#A9D09E", font=("Vivaldi", 14))
        l.place(x = 680, y = 5, width=250, height=50)
        l1 = Label(root, text='꧁ Welcome to LEO - AI Virtual Assistant Application! ꧂', fg='black',bg="#A9D09E", font=("Vivaldi", 14))
        l1.place(x = 70, y = 5, width=510, height=50)
        l1 = Label(root, text='Contact Us:', fg='black', font=("Vivaldi", 15))
        l1.place(x = 10, y = 497, width=100, height=20)
        
        # FOOTER SETTING 
        micro = Button(self, text="Micro",command = ham_main,width=10,fg="white", bg="#009999",bd=3)
        micro.pack(side=RIGHT, padx=11, pady=10)
        clear = Button(self,text="Refresh",command = r_set,width=10,fg="white", bg="#009999",bd=3)
        clear.pack(side=RIGHT,padx=11, pady=10)
        Infor = Button(self,text="About Us",command = info,width=12,fg="white", bg="#009999",bd=3)
        Infor.pack(side=RIGHT,padx=11, pady=10)
        
        # Facebook icon button
        nmn= (Image.open("FYP - AI Virtual Assistant//facebook.png"))
        resized_image= nmn.resize((30,30), Image.Resampling.LANCZOS)
        nmn_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(self, image=nmn_1)
        label1.image = nmn_1
        b = tk.Button(self, image=nmn_1, command=fb_click)
        b.place(x=120, y=488)
        
        # Instagram icon button
        nmn= (Image.open("FYP - AI Virtual Assistant//Instagram.png"))
        resized_image= nmn.resize((30,30), Image.Resampling.LANCZOS)
        nmn_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(self, image=nmn_1)
        label1.image = nmn_1
        b = tk.Button(self, image=nmn_1, command=insta_click)
        b.place(x=155, y=488)
        
        # Whatsapp icon button
        nmn= (Image.open("FYP - AI Virtual Assistant//whatsapp.png"))
        resized_image= nmn.resize((30,30), Image.Resampling.LANCZOS)
        nmn_1= ImageTk.PhotoImage(resized_image)
        label1 = Label(self, image=nmn_1)
        label1.image = nmn_1
        b = tk.Button(self, image=nmn_1, command=wapp_click)
        b.place(x=190, y=488)
        
        # Micro button
        image3 = Image.open("FYP - AI Virtual Assistant//micro.jpg")
        image_3 = ImageTk.PhotoImage(image3)  
        label = Label(image=image_3)
        label.image = image_3
        label.place(x=528, y=497)
        
        
        self.pack(fill=BOTH, expand=1)   
        Style().configure("TFrame", background="#333")
    
root.geometry("1000x550+140+40")
root.resizable(False, False)
app = Example(root)
root.mainloop()
