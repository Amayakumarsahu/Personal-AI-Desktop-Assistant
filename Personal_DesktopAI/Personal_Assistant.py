import datetime
import os
import time
import webbrowser
import sys
import psutil
import pyautogui
import pyttsx3
import speech_recognition as sr
import screen_brightness_control as sbc  
import numpy as np
import pickle
import json
from tensorflow.keras.models import load_model # type: ignore
from tensorflow.keras.preprocessing.sequence import pad_sequences # type: ignore
import ctypes


with open(r"C:\Users\Amaya\OneDrive\Desktop\VS practice codes\Personal_DesktopAI\intents.json") as file:
    data = json.load(file)

model = load_model(r"C:\Users\Amaya\OneDrive\Desktop\VS practice codes\Personal_DesktopAI\chat_model.h5")

with open(r"C:\Users\Amaya\OneDrive\Desktop\VS practice codes\Personal_DesktopAI\tokenizer.pkl", "rb") as f:
    tokenizer=pickle.load(f)

with open(r"C:\Users\Amaya\OneDrive\Desktop\VS practice codes\Personal_DesktopAI\label_encoder.pkl", "rb") as encoder_file:
    label_encoder=pickle.load(encoder_file)

def initialize_engine():
    engine = pyttsx3.init("sapi5")
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - 50)
    volume = engine.getProperty('volume')
    engine.setProperty('volume', volume + 0.25)
    return engine

def speak(text):
    engine = initialize_engine()
    engine.say(text)
    engine.runAndWait()

def command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)  # To remove ambient noise
        print("Listening.......", end="", flush=True)
        r.pause_threshold = 1.0
        r.phrase_threshold = 0.3
        r.sample_rate = 48000
        r.dynamic_energy_threshold = True
        r.operation_timeout = 5
        r.non_speaking_duration = 0.5
        r.dynamic_energy_adjustment = 2
        r.energy_threshold = 4000
        r.phrase_time_limit = 10
        audio = r.listen(source)
    try:
        print("\r", end="", flush=True)
        print("Recognizing.....", end="", flush=True)
        query = r.recognize_google(audio, language='en-in')
        print("\r", end="", flush=True)
        print(f"User said: {query}\n")
    except Exception as e:
        speak("Say that again please...")
        return "None"
    return query

def cal_day():
    day = datetime.datetime.today().weekday() + 1
    day_dict = {
        1: "Monday",
        2: "Tuesday",
        3: "Wednesday",
        4: "Thursday",
        5: "Friday",
        6: "Saturday",
        7: "Sunday",
    }
    if day in day_dict.keys():
        day_of_weeks = day_dict[day]
        print(day_of_weeks)
    return day_of_weeks

def wishme():
    hour = int(datetime.datetime.now().hour)
    t = time.strftime("%I:%M:%p")
    day = cal_day()
    if (hour >= 0) and (hour <= 12) and ('AM' in t):
        speak(f"Good Morning Boss, it's {day} and the time is {t}")
    if (hour >= 12) and (hour <= 16) and ('PM' in t):
        speak(f"Good Afternoon Boss, it's {day} and the time is {t}")
    if (hour >= 16) and ('PM' in t):
        speak(f"Good Evening Boss, it's {day} and the time is {t}")

def social_media(command):
    if 'facebook' in command:
        speak("Opening your Facebook...")
        webbrowser.open("https://www.facebook.com/")
    elif 'discord' in command:
        speak("Opening your Discord server...")
        webbrowser.open("https://www.discord.com/")
    elif 'whatsapp' in command:
        speak("Opening WhatsApp...")
        webbrowser.open("https://web.whatsapp.com/")
    elif 'instagram' in command:
        speak("Opening Instagram....")
        webbrowser.open("https://www.instagram.com")
    elif 'youtube' in command:
        speak("Opening YouTube.......")
        webbrowser.open("https://www.youtube.com")
    elif 'twitter' in command:
        speak("Opening Twitter....")
        webbrowser.open("https://www.twitter.com")
    elif 'linkedin' in command:
        speak("Opening LinkedIn....")
        webbrowser.open("https://www.linkedin.com")
    elif 'github' in command:
        speak("Opening GitHub....")
        webbrowser.open("https://github.com/")

def close_social_media(command):
    if 'facebook' in command:
        speak("Closing Facebook...")
        os.system('taskkill /f /im msedge.exe') 
    elif 'discord' in command:
        speak("Closing Discord...")
        os.system('taskkill /f /im msedge.exe')  
    elif 'whatsapp' in command:
        speak("Closing WhatsApp...")
        os.system('taskkill /f /im msedge.exe')  
    elif 'instagram' in command:
        speak("Closing Instagram...")
        os.system('taskkill /f /im msedge.exe') 
    elif 'youtube' in command:
        speak("Closing YouTube...")
        os.system('taskkill /f /im msedge.exe')  
    elif 'twitter' in command:
        speak("Closing Twitter...")
        os.system('taskkill /f /im msedge.exe') 
    elif 'linkedin' in command:
        speak("Closing LinkedIn...")
        os.system('taskkill /f /im msedge.exe') 
    elif 'github' in command:
        speak("Closing GitHub...")
        os.system('taskkill /f /im msedge.exe')

def openApp(command):
    if "calculator" in command:
        speak("Opening Calculator....")
        os.startfile('C:\\Windows\\System32\\calc.exe')
    elif "notepad" in command:
        speak("Opening Notepad....")
        os.startfile('C:\\Windows\\System32\\notepad.exe')
    elif "paint" in command:
        speak("Opening Paint....")
        os.startfile('C:\\Windows\\System32\\mspaint.exe')
    elif "word" in command or "microsoft word" in command:
        speak("Opening Microsoft Word....")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')
    elif "excel" in command or "microsoft excel" in command:
        speak("Opening Microsoft Excel....")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE')
    elif "powerpoint" in command or "microsoft powerpoint" in command:
        speak("Opening Microsoft PowerPoint....")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE')
    elif "pycharm" in command or "py charm" in command:
        speak("Opening PyCharm....")
        os.startfile(r"C:\Users\Public\Desktop\PyCharm Community Edition 2024.1.4.lnk") 
    elif "vs code" in command or "visual studio code" in command:
        speak("Opening Visual Studio Code....")
        os.startfile(r"C:\Users\Amaya\OneDrive\Desktop\Visual Studio Code.lnk")  

def closeApp(command):
    if "calculator" in command:
        speak("Closing Calculator....")
        os.system('taskkill /f /im calc.exe')
    elif "notepad" in command:
        speak("Closing Notepad....")
        os.system('taskkill /f /im notepad.exe')
    elif "paint" in command:
        speak("Closing Paint....")
        os.system('taskkill /f /im mspaint.exe')
    elif "word" in command or "microsoft word" in command:
        speak("Closing Microsoft Word....")
        os.system('taskkill /f /im WINWORD.EXE')
    elif "excel" in command or "microsoft excel" in command:
        speak("Closing Microsoft Excel....")
        os.system('taskkill /f /im EXCEL.EXE')
    elif "powerpoint" in command or "microsoft powerpoint" in command:
        speak("Closing Microsoft PowerPoint....")
        os.system('taskkill /f /im POWERPNT.EXE')
    elif "whatsapp" in command:
        speak("Closing WhatsApp Desktop....")
        os.system('taskkill /f /im WhatsApp.exe')
    elif "pycharm" in command or "py charm" in command:
        speak("Closing PyCharm....")
        os.system('taskkill /f /im pycharm64.exe')
    elif "vs code" in command or "visual studio code" in command:
        speak("Closing Visual Studio Code....")
        os.system('taskkill /f /im Code.exe')

def control_brightness(command):
    current_brightness = sbc.get_brightness()
    if "brightness up" in command or "increase brightness" in command:
        new_brightness = min(current_brightness[0] + 10, 100)  
        sbc.set_brightness(new_brightness)
        speak(f"Brightness increased to {new_brightness}%")
    elif "brightness down" in command or "decrease brightness" in command:
        new_brightness = max(current_brightness[0] - 10, 0)  
        sbc.set_brightness(new_brightness)
        speak(f"Brightness decreased to {new_brightness}%")

def condition():
    usage = str(psutil.cpu_percent())
    speak(f"CPU is at {usage} percentage")
    battery = psutil.sensors_battery()
    percentage = battery.percent
    speak(f"Boss our system have {percentage} percentage battery")

    if percentage>=80:
        speak("Boss we could have enough charging to continue our recording")
    elif percentage>=40 and percentage<=75:
        speak("Boss we should connect our system to charging point to charge our battery")
    else:
        speak("Boss we have very low power, please connect to charging otherwise recording should be off...")

def browsing(query):
    # Path to Microsoft Edge executable
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    if "google" in query:
        speak("Boss, what should I search on Google?")
        search_query = command().lower()
        if search_query != "none":
            speak(f"Searching for {search_query} on Google using Microsoft Edge...")
            try:
                # Open Microsoft Edge with Google search query
                os.system(f'"{edge_path}" https://www.google.com/search?q={search_query}')
            except Exception as e:
                speak(f"Sorry, I couldn't open Microsoft Edge. Error: {str(e)}")
        else:
            speak("I didn't catch that. Please try again.")

    elif "edge" in query or "microsoft edge" in query:
        speak("Boss, what should I search on Microsoft Edge?")
        search_query = command().lower()
        if search_query != "none":
            speak(f"Searching for {search_query} on Microsoft Edge...")
            try:
                # Open Microsoft Edge with Bing search query
                os.system(f'"{edge_path}" https://www.bing.com/search?q={search_query}')
            except Exception as e:
                speak(f"Sorry, I couldn't open Microsoft Edge. Error: {str(e)}")
        else:
            speak("I didn't catch that. Please try again.")

def shutdown():
    speak("Shutting down the laptop. Goodbye!")
    os.system("shutdown /s /t 1")  # Windows command to shut down

def restart():
    speak("Restarting the laptop. Be right back!")
    os.system("shutdown /r /t 1")  # Windows command to restart

def sleep_mode():
    speak("Putting the laptop into sleep mode. See you soon!")
    ctypes.windll.PowrProf.SetSuspendState(0)


listening_paused = False

def main():
    wishme()
    global listening_paused
    while True:
        if listening_paused:
            
            query = command().lower()
            if "resume listening" in query:
                listening_paused = False
                speak("Resuming listening. How can I assist you?")
            continue  

        query = command().lower()

        if "pause listening" in query:
                listening_paused = True
                speak("Pausing listening. Say 'Resume listening' when you need me again.")
                continue  
        # query = command().lower()
        # query = input("Enter your command --> ").lower()

        if ('open facebook' in query) or ('open discord' in query) or ('open whatsapp' in query) or ('open instagram' in query) or ('open youtube' in query) or ('open twitter' in query) or ('open linkedin' in query) or ('open github' in query):
            social_media(query)
        
        elif ("close facebook" in query) or ("close discord" in query) or ("close whatsapp" in query) or ("close instagram" in query) or ("close youtube" in query) or ("close twitter" in query) or ("close linkedin" in query) or ("close github" in query):
            close_social_media(query)

        elif ("volume up" in query) or ("increase volume" in query):
            pyautogui.press("volumeup")
            speak("Got it! Volume increased.")

        elif ("volume down" in query) or ("decrease volume" in query):
            pyautogui.press("volumedown")
            speak("Alright, volume decreased.")

        elif ("volume mute" in query) or ("mute the sound" in query):
            pyautogui.press("volumemute")
            speak("Volume muted. Let me know if you need sound back!")

        elif ("brightness up" in query) or ("increase brightness" in query):
            control_brightness(query)

        elif ("brightness down" in query) or ("decrease brightness" in query):
            control_brightness(query)

        elif ("open calculator" in query) or ("open notepad" in query) or ("open paint" in query) or ("open word" in query) or ("open excel" in query) or ("open powerpoint" in query) or ("open pycharm" in query) or ("open vs code" in query):
            openApp(query)

        elif ("close calculator" in query) or ("close notepad" in query) or ("close paint" in query) or ("close word" in query) or ("close excel" in query) or ("close powerpoint" in query) or ("close whatsapp" in query) or ("close pycharm" in query) or ("close vs code" in query):
            closeApp(query)

        elif "shut down" in query:
            shutdown()

        elif "restart" in query:
            restart()

        elif "sleep mode" in query:
            sleep_mode()

        elif ("what" in query) or ("who" in query) or ("how" in query) or ("hi" in query) or ("thanks" in query) or ("hello" in query):
            padded_sequences = pad_sequences(tokenizer.texts_to_sequences([query]), maxlen=20, truncating='post')
            result = model.predict(padded_sequences)
            tag = label_encoder.inverse_transform([np.argmax(result)])

            for intent in data['intents']:
                if intent['tag'] == tag:
                    # Select a random response from the list
                    response = np.random.choice(intent['responses'])
                    # Modify the response dynamically
                    if "greeting" in tag:
                        response = f"{response} How can I assist you today?"
                    elif "thanks" in tag:
                        response = f"{response} Anytime, Boss!"
                    elif "farewell" in tag:
                        response = f"{response} Take care and see you soon!"
                    elif "jokes" in tag:
                        response = f"Here's one for you: {response}"
                    speak(response)
                    break

        elif ("system condition" in query) or ("condition of the system" in query):
            speak("Checking the system condition...")
            condition()

        elif ("open google" in query) or ("open  microsoft edge" in query) or ("open edge" in query):
            browsing(query)

        elif "exit" in query:
            speak("Thank you, Boss! Let me know if you need my help again. see you soon...... !")
            sys.exit()

if __name__ == "__main__":
    main()