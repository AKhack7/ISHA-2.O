import tkinter as tk
from tkinter import messagebox, scrolledtext
import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
import random
import os
import requests
import subprocess

# Text-to-speech engine setup
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 175)

def speak(text):
    chat_box.insert(tk.END, "Isha: " + text + "\n")
    chat_box.see(tk.END)
    engine.say(text)
    engine.runAndWait()

def get_time():
    time_now = datetime.datetime.now().strftime("%I:%M %p")
    speak("Current time is " + time_now)

def get_date():
    date_today = datetime.date.today().strftime("%B %d, %Y")
    speak("Today's date is " + date_today)

def solve_math(expression):
    try:
        result = eval(expression)
        speak("The result is " + str(result))
    except:
        speak("Sorry, I could not solve that.")

def show_settings_popup():
    settings = [
        "Display setting (01)", "Sound setting (03)", "Notification & action setting (07)", "Focus assist setting (08)",
        "Power & sleep setting (04)", "Storage setting (05)", "Tablet setting (03)", "Multitasking setting (088)",
        "Projecting to this pc setting (099)", "Shared experiences setting (076)", "System components setting (098)",
        "Clipboard setting (054)", "Remote desktop setting (00)", "Optional features setting (021)", "About setting (007)",
        "System setting (0022)", "Devices setting (0033)", "Mobile devices setting (0044)", "Network & internet setting (0055)",
        "Personalization setting (0066)", "Apps setting (0099)", "Account setting (0088)", "Time & language setting (0010)",
        "Gaming setting (0009)", "Ease of access setting (0080)", "Privacy setting (0076)", "Updated & security (0087)"
    ]
    popup = tk.Toplevel(window)
    popup.title("All Settings")
    popup.geometry("500x400")
    tk.Label(popup, text="Settings List:", font=('Arial', 14)).pack()
    scrolled = scrolledtext.ScrolledText(popup, wrap=tk.WORD, width=50, height=20)
    scrolled.pack(padx=10, pady=10)
    for item in settings:
        scrolled.insert(tk.END, item + "\n")

def show_apps_popup():
    apps = [
        "Alarms & Clock", "Calculate", "Calendar", "Camera", "Copilot", "Cortana", "Game Bar", "Groove music", "Mail",
        "Maps", "Microsoft edge", "Microsoft solitaire collection", "Microsoft Store", "Mixed reality portal", "Movies & tv",
        "Office", "Onedrive", "Onenote", "Outlook", "Outlook (classic)", "Paint", "Paint 3D", "Phone link", "Power point",
        "Settings", "Skype", "Snip & sketch", "Stucky note", "Tips", "Voice recorder", "Whether", "Windows backup",
        "Windows security", "Word", "Xbox", "About your pc", "notepad", "cmd", "Excle", "Control panel", "File explorer"
    ]
    popup = tk.Toplevel(window)
    popup.title("All Apps")
    popup.geometry("500x400")
    tk.Label(popup, text="App List:", font=('Arial', 14)).pack()
    scrolled = scrolledtext.ScrolledText(popup, wrap=tk.WORD, width=50, height=20)
    scrolled.pack(padx=10, pady=10)
    for app in apps:
        scrolled.insert(tk.END, app + "\n")

def open_app(app_name):
    app_paths = {
        "alarms & clock": "start ms-clock:",
        "calculate": "calc",
        "calendar": "start outlookcal:",
        "camera": "start microsoft.windows.camera:",
        "copilot": "start ms-copilot:",
        "cortana": "start ms-cortana:",
        "game bar": "start xbox-gamebar:",
        "groove music": "start mswindowsmusic:",
        "mail": "start outlookmail:",
        "maps": "start bingmaps:",
        "microsoft edge": "start msedge:",
        "microsoft solitaire collection": "start microsoft.microsoftsolitairecollection:",
        "microsoft store": "start ms-windows-store:",
        "mixed reality portal": "start ms-holographicfirstlaunch:",
        "movies & tv": "start mswindowsvideo:",
        "office": "start ms-officeapp:",
        "onedrive": "start onedrive:",
        "onenote": "start onenote:",
        "outlook": "start outlook:",
        "outlook (classic)": "start outlook:",
        "paint": "mspaint",
        "paint 3d": "start ms-paint:",
        "phone link": "start ms-phone:",
        "power point": "powerpnt",
        "settings": "start ms-settings:",
        "skype": "start skype:",
        "snip & sketch": "start ms-screenclip:",
        "stucky note": "start stikynot:",
        "tips": "start ms-get-started:",
        "voice recorder": "start ms-voicerecorder:",
        "whether": "start bingweather:",
        "windows backup": "start ms-settings:backup",
        "windows security": "start windowsdefender:",
        "word": "winword",
        "xbox": "start xbox:",
        "about your pc": "start ms-settings:about",
        "notepad": "notepad",
        "cmd": "cmd",
        "excle": "excel",
        "control panel": "control",
        "file explorer": "explorer"
    }
    app_name = app_name.lower()
    if app_name in app_paths:
        try:
            os.system(app_paths[app_name])
            speak("Opening " + app_name)
        except:
            speak("Could not open " + app_name)
    else:
        speak("App not found")

def get_weather():
    api_key = "YOUR_API_KEY_HERE"
    url = f"https://api.openweathermap.org/data/2.5/weather?q=Ahmedabad&appid={api_key}&units=metric"
    try:
        response = requests.get(url).json()
        temp = response['main']['temp']
        desc = response['weather'][0]['description']
        speak(f"The current temperature in Ahmedabad is {temp} degree Celsius with {desc}")
    except:
        speak("Failed to retrieve weather information")

def handle_command(command):
    command = command.lower()
    chat_box.insert(tk.END, "You: " + command + "\n")
    chat_box.see(tk.END)

    if any(op in command for op in ["+", "-", "*", "/"]):
        solve_math(command)
    elif "time" in command or "samaye" in command:
        get_time()
    elif "date" in command or "aaj" in command:
        get_date()
    elif "solve" in command or "calculate" in command:
        speak("Please type the expression to solve.")
    elif "about all setting" in command:
        show_settings_popup()
    elif "about all app" in command:
        show_apps_popup()
    elif "open" in command:
        parts = command.split()
        for p in parts:
            if p.isdigit():
                open_setting_by_code(p)
                return
        app_name = command.replace("open", "").strip()
        if app_name in ["google", "youtube", "whatsapp", "facebook", "instagram"]:
            urls = {
                "google": "https://www.google.com",
                "youtube": "https://www.youtube.com",
                "whatsapp": "https://web.whatsapp.com",
                "facebook": "https://www.facebook.com",
                "instagram": "https://www.instagram.com"
            }
            webbrowser.open(urls[app_name])
            speak("Opening " + app_name)
        else:
            open_app(app_name)
    elif "play song" in command or "play music" in command:
        songs = [
            "https://youtu.be/-Me4lYdECn0?si=ymrom6m8nEQpBLFe",
            "https://youtu.be/ydk-JTPln9I?si=Fo-AxFx7VW3XT5Jc",
            "https://youtu.be/YJg1rs0R2sE?si=hPsmwlU2DzOSvweX"
        ]
        webbrowser.open(random.choice(songs))
        speak("Playing music")
    elif "find now" in command:
        speak("What should I search?")
    elif "thank you" in command:
        speak("Welcome")
    elif "good night" in command:
        speak("Shutting down, good night")
        os.system("shutdown /s /t 1")
    elif "weather" in command or "wether" in command or "ahmedabad" in command:
        get_weather()
    else:
        speak("I did not understand. Try again.")

def voice_input():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        speak("Listening...")
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio)
            handle_command(query)
        except sr.UnknownValueError:
            speak("Sorry, I did not understand.")

def send_text(event=None):
    query = input_box.get()
    input_box.delete(0, tk.END)
    handle_command(query)

# GUI setup
window = tk.Tk()
window.title("ISHA Assistant")
window.geometry("600x500")

chat_box = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=70, height=25)
chat_box.pack(padx=10, pady=10)

frame = tk.Frame(window)
frame.pack()

input_box = tk.Entry(frame, width=50)
input_box.grid(row=0, column=0, padx=5)
input_box.bind('<Return>', send_text)

send_button = tk.Button(frame, text="Send", command=send_text)
send_button.grid(row=0, column=1, padx=5)

voice_button = tk.Button(window, text="🎤 Voice", bg="red", fg="white", command=voice_input)
voice_button.pack(pady=10)

speak("Hello, I am ISHA. INTELLIGENT SYSTEM FOR HUMAN ASSISTANCE.")

window.mainloop()
