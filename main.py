import pyttsx3
import speech_recognition as sr
import datetime
import time
import webbrowser
import pyautogui
import subprocess

import json
import pickle
from tensorflow import keras
import random
from tensorflow.keras.preprocessing.sequence import pad_sequences # type: ignore
import numpy as np
from tensorflow.keras.models import load_model # type: ignore
import psutil 
#from elevenlabs import generate, play, set_api_key # type: ignore
#from elevenlabs import set_api_key # type: ignore
#from api_key import api_key_data # type: ignore

#set_api_key(api_key_data)

#def engine_talk(query):
 #   audio = generate(
 #       text = query,
 #       voice = 'Grace',
 #       model = 'eleven_monolingual_v1'
 #  )
 #   play(audio)

with open('intents.json') as file:
    data = json.load(file)

model  = load_model('chat_model.h5')

# Saving tokenizer
with open("tokenizer.pkl", "rb") as f:
    tokenizer = pickle.load(f)

# Saving label encoder
with open("label_encoder.pkl", "rb") as encoder_file:
    label_encoder = pickle.load(encoder_file)

# Initialize Text-to-Speech Engine
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)  # Use a female voice if available
engine.setProperty("rate", 170)  # Slightly faster speech rate
engine.setProperty("volume", 1.0)  # Max volume

def speak(text):
    """Speaks the given text if it's valid."""
    if text.strip() and text.lower() != "none":  # Ensures valid speech
        engine.say(text)
        engine.runAndWait()

def command():
    """Listens and recognizes speech quickly."""
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)  # Reduce noise adjustment time
        print("Listening...", end=" ", flush=True)

        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=4)  # Reduce time limits
            print("Recognizing...")

            query = r.recognize_google(audio, language="en-IN").strip().lower()  
            if query == "":
                return None  # Return None explicitly if empty
            
            
            return query

        except sr.UnknownValueError:
            print("Could not understand, please speak clearly.")
        except sr.RequestError:
            print("Speech Recognition service unavailable.")
        except sr.WaitTimeoutError:
            print("No speech detected. Try again.")

    return None  # Explicitly return None when no input is detected

def cal_day():
    day = datetime.datetime.today().weekday() + 1
    day_dict = {
        1: 'monday',
        2: 'tuesday',
        3: 'wednesday',
        4: 'thursday',
        5: 'friday',
        6: 'saturday',
        7: 'sunday'
    }
    return day_dict.get(day, "unknown")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    t = time.strftime("%I:%M %p")
    day = cal_day()
    
    if (hour >= 0 and hour < 12 and 'AM' in t):
        speak(f"Good Morning, Akhil! It's {day}, and the time is {t}")
    elif (hour >= 12 and hour < 18 and 'PM' in t):
        speak(f"Good Afternoon, Akhil! It's {day}, and the time is {t}")
    else:
        speak(f"Good Evening, Akhil! It's {day}, and the time is {t}")

def social_media(command):
    if 'facebook' in command:
        speak("Opening Facebook")
        webbrowser.open("https://www.facebook.com")
    elif 'discord' in command:
        speak("Opening Discord")
        webbrowser.open("https://www.discord.com")
    elif 'whatsapp' in command:
        speak("Opening WhatsApp")
        try:
             subprocess.run(['powershell', 'Start-Process', 'whatsapp:'], check=True)
        except FileNotFoundError:
            speak("WhatsApp not found on this device")       
    elif 'instagram' in command:
        speak("Opening Instagram")
        webbrowser.open("https://www.instagram.com")
    else:
        speak("Sorry, I can't open that")    

def schedule(command):
    if 'schedule' in command or 'time table' in command:
        day = cal_day().lower()
        speak("Boss, today's schedule is as follows:")
        week_schedule = {
            'monday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'tuesday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'wednesday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'thursday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'friday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'saturday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM.",
            'sunday': "W3 Schools from 9:00 AM to 11:00 AM, HackerRank from 1:00 PM to 3:00 PM, Badminton from 4:00 PM to 6:00 PM, Project from 7:00 PM to 9:00 PM."
        }

        speak(week_schedule.get(day, "Sorry, I can't show your schedule."))

def closeApp(command):
    if 'close calculator' in command:
        speak("Closing calculator")
        try:
            subprocess.run(["powershell", "-Command", "Stop-Process -Name CalculatorApp -Force"], check=True)
        except subprocess.CalledProcessError:
            speak("Calculator not found on this device")
    if 'close notepad' in command:
        speak("Closing Notepad")
        try:
            subprocess.run(["powershell", "-Command", "Stop-Process -Name notepad -Force"], check=True)
        except subprocess.CalledProcessError:
            speak("Notepad not found on this device")
    if 'close paint' in command:
        speak("Closing paint")
        try:
            subprocess.run(["powershell", "-Command", "Stop-Process -Name mspaint -Force"], check=True)
        except subprocess.CalledProcessError:
            speak("paint not found on this device")
    if 'close word' in command:
        speak("Closing word")
        try:
            subprocess.run(["powershell", "-Command", "Stop-Process -Name WINWORD -Force"], check=True)
        except subprocess.CalledProcessError:
            speak("word not found on this device")
    if 'close excel' in command:
        speak("Closing excel")
        try:
            subprocess.run(["powershell", "-Command", "Stop-Process -Name EXCEL -Force"], check=True)
        except subprocess.CalledProcessError:
            speak("excel not found on this device")

def openApp(command):
    if 'open calculator' in command:
        speak("Opening calculator")
        try:
            subprocess.run(["powershell", "-Command", "Start-Process calc"], check=True)
        except subprocess.CalledProcessError:
            speak("Calculator not found on this device")
    if 'open notepad' in command:
        speak("Opening Notepad")
        try:
            subprocess.run(["powershell", "-Command", "Start-Process notepad"], check=True)
        except subprocess.CalledProcessError:
            speak("Notepad not found on this device")
    if 'open paint' in command:
        speak("Opening paint")
        try:
            subprocess.run(["powershell", "-Command", "Start-Process paint"], check=True)
        except subprocess.CalledProcessError:
            speak("paint not found on this device")
    if 'open word' in command:
        speak("Opening word")
        try:
            subprocess.run(["powershell", "-Command", "Start-Process winword"], check=True)
        except subprocess.CalledProcessError:
            speak("word not found on this device")
    if 'open excel' in command:
        speak("Opening excel")
        try:
            subprocess.run(["powershell", "-Command", "Start-Process excel"], check=True)
        except subprocess.CalledProcessError:
            speak("excel not found on this device")
        

def browsing(command):
    if 'open google' in command:
        speak("Opening Google")
        speak("Boss, what do you want to search on Google?")
        s = input("What do you want to search for? ").lower()
        search_url = f"https://www.google.com/search?q={s.replace(' ', '+')}"
        subprocess.run(["powershell", "-Command", f"Start-Process -FilePath 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' -ArgumentList '{search_url}'"], check=True)

    elif 'open edge' in command:
        speak("Opening Edge")
        speak("Boss, what do you want to search on Bing?")
        s = input("What do you want to search for? ").lower()
        search_url = f"https://www.bing.com/search?q={s.replace(' ', '+')}"
        subprocess.run(["powershell", "-Command", f"Start-Process -FilePath 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe' -ArgumentList '{search_url}'"], check=True)

    elif 'open firefox' in command:
        speak("Opening Firefox")
        speak("Boss, what do you want to search on Firefox?")
        s = input("What do you want to search for? ").lower()
        search_url = f"https://www.google.com/search?q={s.replace(' ', '+')}"
        subprocess.run(["powershell", "-Command", f"Start-Process -FilePath 'C:\\Program Files\\Mozilla Firefox\\firefox.exe' -ArgumentList '{search_url}'"], check=True)

            
def condition():
    usage = str(psutil.cpu_percent(interval=1))
    speak(f"CPU usage is at {usage}%")
    battery = psutil.sensors_battery()
    percentage = battery.percent
    speak(f"Boss, Battery percentage is at {percentage}%")
    
    if percentage < 30:
        speak("Boss, Battery is low. Please connect the charger")
    elif  30 <= percentage <= 80:
        speak("Boss, Battery is at moderate level")
    else:
        speak("Boss, Battery is at high level, you can continue working")

if __name__ == "__main__":
    wishMe()
    #engine_talk("Allow me to introduce myself I am Jarvis, your personal assistant. How can I help you today?")
    
    while True:
        query = command()
        #query = imput("Enter your query: ")
        if query:  # If query is not None or empty
            print(f"You said: {query}")
            speak(f"You said: {query}")  
            if "hello" in query or "hi" in query:
                speak("Hello Akhil! How can I help you today?")
            elif "how are you" in query:
                speak("I'm doing great, Akhil! Thanks for asking.")
            elif "thank you" in query or "thanks" in query:
                speak("You're welcome, Akhil!")
            elif 'facebook' in query or 'discord' in query or 'whatsapp' in query or 'instagram' in query:
                social_media(query)
            elif 'schedule' in query or 'time table' in query:
                schedule(query)    
            elif 'volume up' in query or 'increase volume' in query:
                pyautogui.press('volumeup')
                speak("Volume increased")
            elif 'volume down' in query or 'decrease volume' in query:
                pyautogui.press('volumedown')
                speak("Volume decreased")
            elif 'mute' in query or 'volume mute' in query:
                pyautogui.press('volumemute')
                speak("Volume muted")
            elif ('open calculator' in query) or ('open notepad' in query) or ('open paint' in query) or ('open word' in query) or ('open excel' in query):
                openApp(query)    
            elif ('close calculator' in query) or ('close notepad' in query) or ('close paint' in query) or ('close word' in query) or ('close excel' in query):
                closeApp(query)
            elif ("what" in query or "who" in query or "how" in query or "hi" in query or "hello" in query or "thanks" in query or "thank you" in query):
                padded_sequences = pad_sequences(tokenizer.texts_to_sequences([query]), truncating='post', maxlen=20)
                result = model.predict(padded_sequences)
                tag = label_encoder.inverse_transform([np.argmax(result)])
                
                for i in data['intents']:
                    if i['tag'] == tag:
                        speak(np.random.choice(i['responses']))
            elif ("open google" in query) or ("open edge" in query) or ("open firefox" in query):
                browsing(query)
            elif ("system condition" in query) or ("system status" in query) or ("system information" in query):
                speak("Boss, here is your system information")
                condition()
            elif "exit" in query or "stop" in query:
                speak("Goodbye Akhil! Have a great day.")
                break  # Exit loop

