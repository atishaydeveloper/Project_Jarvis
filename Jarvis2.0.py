import speech_recognition as sr
import pyttsx3 
import datetime
import spacy 
import re
import webbrowser
import os
import requests
import subprocess
import pyautogui
import time
import win32com.client
import tkinter as tk
from tkinter import scrolledtext
import json
import sys


#https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



data = '{"key": null}'

def parse_json_with_nulls(json_str):
    json_str = json_str.replace("null", "None")  # Replace `null` for compatibility
    json_str = json_str.replace("None", "null")  # Ensure `None` compatibility
    return eval(json_str) 




nlp = spacy.load(resource_path("en_core_web_sm"))
engine = pyttsx3.init()
reminders = {}
API_KEY = "bfff046149909e34cd76fb5f61c352a9"
def speak(text):
    try:
        engine.say(text)
        engine.runAndWait()
    except RuntimeError:
        print("Speech synthesis is currently busy. Retrying...")
        
        
def listen():
    
    recognizer = sr.Recognizer()
    try:
        with  sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)
            display_message("Please say something...")
            print("Please say something...")
            
            audio_data = recognizer.listen(source)
            print("Processing audio...")
    
            try:
                command = recognizer.recognize_google(audio_data)
                print("You said:", command)
                return command.lower()
            except sr.UnknownValueError:
                print("Sorry, could not understand what you said.")
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
    except OSError as e:
        print("Microphone is not available:",e)    
    return ""
    
def get_weather(city):
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        weather_desc = data["weather"][0]["description"]
        temperature = data["main"]["temp"]
        return f"The weather in {city} is currently {weather_desc} with a temperature of {temperature} degree Celsius."
    else:
        return "Sorry, I couldn't fetch the weather information."
        
def get_intent_and_entities(command):
    if "time" in command:
        intent = "get_time"
    elif "date" in command:
        intent = "get_date"
    elif "lock" in command:
        intent = "lock"
    elif any(greeting in command for greeting in ["hello", "hey", "jarvis"]):
        intent = "greet"
    elif any(exit in command for exit in ["bye","exit","later"] ):
        intent = "exit"
    elif "weather" in command:
        intent = "get_weather"
    elif "open" in command:
        intent = "open_application"
    elif any(assist in command for assist in ["help", "capable", "do","assist"]):
        intent = "assist"
    elif "reminder" in command:
        intent = "set_reminder"
    elif "search" in command:  # New intent for searching
        intent = "search_web"
    elif any(word in command for word in ["calculate", "plus", "minus", "divide"]):
        intent = "calculate"
    else:
        intent = "unknown"
        
    doc = nlp(command)
    entities = {ent.label_: ent.text for ent in doc.ents}
    return intent, entities

def perform_calculation(command):
    try:
        match = re.findall(r"(\d+|\+|\-|\*|\/)",command)
        expression = "".join(match)
        result = eval(expression)
        speak(f"The result is {result}")
    except Exception:
        speak("I'm sorry, I couldn't perform the calculation.")
        
def display_message(message):
    output_text.insert(tk.END,message+"\n")
    output_text.see(tk.END)


def respond_to_command(intent, entities,command):
    if intent == "get_time":
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The time is {current_time}")
        display_message(f"The time is {current_time}")
    elif intent == "get_date":
        current_date = datetime.datetime.now().strftime("%B %d, %Y")
        speak(f"Today's date is {current_date}")
        display_message(f"Today's date is {current_date}")
    elif intent == "greet":
        speak("Hello, how can I assist you today?")
    elif intent == "exit":
        speak("Goodbye! Have a nice day.")
        return True
    elif intent == "assist":
        display_message("1: setting reminders")
        display_message("2: searching on google")
        display_message("3: opening browser, calculater or notepad")
        display_message("4: calculate mathematical expressions")
        display_message("5: Today's weather, date, or even time")
        speak("I can help you with a variety of tasks, such as setting reminders, searching on google, opening browser, calculater or notepad, and even can calculate mathematical expressions. Along with it I can give info on Today's weather, date, or even time")
        
    elif intent == "lock":
        speak("Locking the computer.")
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('^%{DEL}')
    elif intent == "get_weather":
        city = entities.get("GPE", "London")
        weather_info = get_weather(city)
        speak(weather_info)
        display_message(weather_info)
    elif intent == "open_application":
        app_name = command.split("open ", 1)[1]
        if "browser" in app_name:
            speak("Opening browser.")
            webbrowser.open("http://www.google.com")
        elif "notepad" in app_name:
            speak("Opening Notepad.")
            os.startfile("notepad.exe")
        elif "calculator" in app_name:
            speak("Opening Calculator.")
            os.startfile("calc.exe")
        elif "vs code" in app_name:
            speak("Opening VS code.")
            file_path = "C:\\Users\\ASUS\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Visual Studio Code\\Visual Studio Code.exe"
            if os.path.exists(resource_path(file_path)):
                subprocess.Popen([resource_path(file_path)])
                
            else:
                speak("VScode not found.")
            
        elif "game" in app_name:
            speak("Opening ForzaHorizon5.")
            if os.path.exists(resource_path("D:\\Forza Horizon 5\\ForzaHorizon5.exe")):
                subprocess.Popen(resource_path("D:\\Forza Horizon 5\\ForzaHorizon5.exe"))
            else:
                speak("Sorry, I couldn't find the game.")
        else:
            speak("Sorry, I couldn't find that application.")
            
    elif intent == "set_reminder":
        reminder = command.split("reminder ", 1)[1]
        reminders[len(reminders) + 1] = reminder
        speak(f"Reminder set for: {reminder}.")
        display_message(f"Reminder set for: {reminder}.")
    
    elif intent == "search_web" and "search " in command:
        search_query = command.split("search ", 1)[1]  
        url = f"https://www.google.com/search?q={search_query}"
        webbrowser.open(url) 
        speak(f"Searching for {search_query} on Google.")
                
    elif intent == "calculate":
        perform_calculation(command)   
        display_message(command)
    else:
        speak("Sorry, I didn't understand that. Can you please rephrase?")
    return False

def on_listen_button_click():
    command = listen()
    if command:
        intent,entities = get_intent_and_entities(command)
        if respond_to_command(intent,entities,command):
            root.destroy()
            

root = tk.Tk()
root.title("Jarvis: Voice Assistant")
root.geometry("600x400")

output_text = scrolledtext.ScrolledText(root, wrap = tk.WORD, width = 70, height = 15, font = ("Arial",10))
output_text.pack(padx=10,pady=10)

listen_button = tk.Button(root, text = "Listen", command = on_listen_button_click, font=("Arial",12))
listen_button.pack(padx=10,pady=10)
root.mainloop()