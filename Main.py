# Isha AI Assistant - Voice Only Final Version (Cord7)

import pyttsx3
import speech_recognition as sr
import pyautogui
import os
import webbrowser
import requests
from deep_translator import GoogleTranslator
import threading
from openai import OpenAI
import winreg
import tkinter as tk
from tkinter import scrolledtext
import subprocess
import random
import datetime
import json

# ----------------------------- CONFIG ----------------------------- #
WAKE_WORD = "isha"
API_KEY = "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"
LANGUAGE_TRANSLATE_TO = "hi"
USE_ONLINE = True
SONG_FOLDER = os.path.expanduser("~/Music")
OPENWEATHER_API_KEY = "YOUR_OPENWEATHER_API_KEY"
CITY = "Ahmedabad"

SONG_LINKS = [
    "https://youtu.be/s4KzN7mW8T8?si=-xB_gzSjGfpDSmZM",
    "https://youtu.be/MGwiCtsbB6k?si=5_xcM__lAOJMFc9n",
    "https://youtu.be/uFbayWnLGxs?si=mX5geBfOevjY1vso",
    "https://youtu.be/aHuuaIAS_U4?si=F-IaqgPpJyHVnFoA",
    "https://youtu.be/TkAiQJzctFY?si=a4949Ki95Hu_pE36",
]

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=API_KEY,
)

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 175)

shortcut_states = {}
recognizer = sr.Recognizer()
installed_apps = set()

window = tk.Tk()
window.title("Isha")
window.geometry("600x500")
window.resizable(False, False)
window.configure(bg="lightblue")

log_box = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=70, height=20, font=("Arial", 10, "italic"), fg="black", bg="lightblue", borderwidth=2, relief="solid")
log_box.pack(pady=10)

status_label = tk.Label(window, text="Status: Ready", fg="green", bg="lightblue")
status_label.pack()

entry = tk.Entry(window, width=50)
entry.pack(pady=5)
entry.insert(0, "Type your command here")


def respond(text):
    log_box.insert(tk.END, f"Isha:{text}\n")


def wish_me():
    current_hour = datetime.datetime.now().hour
    if 5 <= current_hour < 12:
        speak("Good morning")
    elif 12 <= current_hour < 17:
        speak("Good afternoon")
    elif 17 <= current_hour < 21:
        speak("Good evening")
    else:
        speak("Good night")

def log(message):
    if message.startswith("You"):
        print(message)
    log_box.insert(tk.END, message + "\n")
    log_box.yview(tk.END)

def speak(text):
    log_box.insert(tk.END, "Isha: " + text + "\n")
    log_box.yview(tk.END)
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=5)
            query = recognizer.recognize_google(audio)
            log("You (voice): " + query)
            return query.lower()
        except:
            return ""

def is_connected():
    try:
        requests.get("https://www.google.com", timeout=3)
        return True
    except:
        return False

def toggle_shortcut(key_combo, description):
    pyautogui.hotkey(*key_combo)
    shortcut_states[description] = not shortcut_states.get(description, False)

shortcut_map = {
    f"ctrl {chr(c)}": (["ctrl", chr(c)], f"ctrl+{chr(c)}") for c in range(ord('a'), ord('z') + 1)
}

def get_installed_apps():
    global installed_apps
    apps = set()
    uninstall_keys = [
        r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
        r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    ]
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for key_path in uninstall_keys:
            try:
                reg_key = winreg.OpenKey(root, key_path)
                for i in range(0, winreg.QueryInfoKey(reg_key)[0]):
                    subkey = winreg.EnumKey(reg_key, i)
                    app_key = winreg.OpenKey(reg_key, subkey)
                    try:
                        name = winreg.QueryValueEx(app_key, 'DisplayName')[0]
                        apps.add(name.lower())
                    except FileNotFoundError:
                        continue
            except Exception:
                pass
    installed_apps = apps

def launch_app(name):
    name = name.lower()
    system_apps = {
        "notepad": "notepad",
        "calculator": "calc",
        "paint": "mspaint",
        "command prompt": "cmd",
        "file explorer": "explorer",
        "this pc": "explorer",
        "settings": "start ms-settings:",
        "control panel": "control",
        "excel": "excel",
        "word": "winword"
    }
    website_map = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com"
    }
    if name in system_apps:
        os.system(system_apps[name])
        speak(f"Opening {name}")
        return
    if name in website_map:
        webbrowser.open(website_map[name])
        speak(f"Opening {name} website")
        return
    matched = [app for app in installed_apps if name in app]
    if matched:
        os.startfile(matched[0])
        speak(f"Opening {matched[0]}")
    else:
        webbrowser.open(f"https://www.google.com/search?q={name}")
        speak(f"{name} not found. Searching on browser.")

def launch_setting(query):
    for name, (uri, code) in SETTING_MAP.items():
        if query == code or query in name:
            os.system(f"start {uri}")
            speak(f"Opening {name}")
            return
    speak("Sorry, setting not found.")

SETTING_MAP = {
    "display setting": ("ms-settings:display", "01"),
    "sound setting": ("ms-settings:sound", "03"),
    "notification & action setting": ("ms-settings:notifications", "07"),
    "focus assist setting": ("ms-settings:quiethours", "08"),
    "power & sleep setting": ("ms-settings:powersleep", "04"),
    "storage setting": ("ms-settings:storagesense", "05"),
    "tablet setting": ("ms-settings:tabletmode", "03"),
    "multitasking setting": ("ms-settings:multitasking", "088"),
    "projecting to this pc setting": ("ms-settings:project", "099"),
    "shared experiences setting": ("ms-settings:crossdevice", "076"),
    "system components setting": ("ms-settings:appsfeatures-app", "098"),
    "clipboard setting": ("ms-settings:clipboard", "054"),
    "remote desktop setting": ("ms-settings:remotedesktop", "00"),
    "optional features setting": ("ms-settings:optionalfeatures", "021"),
    "about setting": ("ms-settings:about", "007"),
    "system setting": ("ms-settings:system", "0022"),
    "devices setting": ("ms-settings:devices", "0033"),
    "mobile devices setting": ("ms-settings:mobile-devices", "0044"),
    "network & internet setting": ("ms-settings:network", "0055"),
    "personalization setting": ("ms-settings:personalization", "0066"),
    "apps setting": ("ms-settings:appsfeatures", "0099"),
    "account setting": ("ms-settings:yourinfo", "0088"),
    "time & language setting": ("ms-settings:dateandtime", "0010"),
    "gaming setting": ("ms-settings:gaming", "0009"),
    "ease of access setting": ("ms-settings:easeofaccess", "0080"),
    "privacy setting": ("ms-settings:privacy", "0076"),
    "updated & security": ("ms-settings:windowsupdate", "0087")
}

def show_all_settings_popup():
    popup = tk.Toplevel(window)
    popup.title("All Settings List")
    popup.geometry("600x600")
    popup.configure(bg="white")

    label = tk.Label(popup, text="Windows Settings Shortcut List", font=("Arial", 14, "bold"), bg="white")
    label.pack(pady=10)

    setting_text = ""
    for name, (_, code) in SETTING_MAP.items():
        setting_text += f"{name.title()} ({code})\n"

    text_area = scrolledtext.ScrolledText(popup, wrap=tk.WORD, width=70, height=30, font=("Arial", 10), bg="#f0f0f0")
    text_area.insert(tk.END, setting_text)
    text_area.configure(state='disabled')
    text_area.pack(pady=10)

def show_all_apps_popup():
    app_popup = tk.Toplevel(window)
    app_popup.title("Common PC Apps")
    app_popup.geometry("400x400")
    app_popup.configure(bg="white")

    label = tk.Label(app_popup, text="Installed App Shortcuts", font=("Arial", 14, "bold"), bg="white")
    label.pack(pady=10)

    apps_list = [
        "Alarms & Clock (1)", "Calculator (2)", "Calendar (3)", "Camera (4)",
        "Copilot (5)", "Cortana (6)", "Game Bar (7)", "Groove Music (8)",
        "Mail (9)", "Maps (0)", "Microsoft Edge (00)", "Microsoft Solitaire Collection (01)",
        "Microsoft Store (02)", "Mixed Reality Portal (03)", "Movies & TV (04)",
        "Office (05)", "OneDrive (06)", "OneNote (07)", "Outlook (08)",
        "Outlook (Classic) (09)", "Paint (000)", "Paint 3D (001)"
    ]

    text_area = scrolledtext.ScrolledText(app_popup, wrap=tk.WORD, width=40, height=20, font=("Arial", 10), bg="#f0f0f0")
    text_area.insert(tk.END, "\n".join(apps_list))
    text_area.configure(state='disabled')
    text_area.pack(pady=10)

    setting_text = ""
    for name, (_, code) in SETTING_MAP.items():
        setting_text += f"{name.title()} ({code})\n"

def play_songs():
    link = random.choice(SONG_LINKS)
    webbrowser.open(link)
    speak("Playing a random song for you")

def get_weather():
    try:
        url = f"https://api.openweathermap.org/data/2.5/weather?q={CITY}&appid={OPENWEATHER_API_KEY}&units=metric"
        response = requests.get(url)
        data = response.json()
        desc = data['weather'][0]['description']
        temp = data['main']['temp']
        speak(f"Current weather in {CITY} is {desc} with {temp} degree Celsius")
    except:
        speak("Couldn't fetch the weather details")

def online_answer(prompt):
    try:
        response = client.chat.completions.create(
            model="openai/gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=100
        )
        return response.choices[0].message.content.strip()
    except:
        return "I'm having trouble reaching the internet right now."

def handle_command(command):
    if "about all setting" in command:
        show_all_settings_popup()
    elif "about all app" in command:
        show_all_apps_popup()
    elif any(x in command for x in ["time", "samy"]):
        speak("Current time is " + datetime.datetime.now().strftime("%H:%M"))
    elif any(x in command for x in ["date", "tarik"]):
        speak("Today's date is " + datetime.datetime.now().strftime("%d-%m-%Y"))
    elif any(x in command for x in ["weather", "mausam", "mosam"]):
        get_weather()
    elif command.startswith("open"):
        app_name = command.replace("open", "").strip()
        if app_name.isdigit() or "setting" in app_name:
            launch_setting(app_name)
        else:
            launch_app(app_name)
    elif "play song" in command or "play music" in command:
        play_songs()
    elif any(op in command for op in ["+", "-", "*", "/"]):
        try:
            result = str(eval(command))
            speak("The answer is " + result)
        except:
            speak("Sorry, I couldn't calculate that")
    elif is_connected() and USE_ONLINE:
        response = online_answer(command)
        speak("Isha: " + response)
    else:
        speak("I can't answer that offline.")

### ---------------------- Reply --------------------------- ###

def handle_command(command):
    command = command.lower().strip()
    log(f"Processing command: {command}")

    # Basic greetings and info
    if any(x in command for x in ["time", "samy"]):
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        response = f"Current time is {current_time}"
        speak(response)
        return

    elif any(x in command for x in ["date", "tarik"]):
        today = datetime.datetime.now().strftime("%d %B %Y")
        response = f"Today's date is {today}"
        speak(response)
        return

    elif command in ["hi", "hello", "hey"]:
        response = "Hello! How can I assist you today?"
        speak(response)
        return

    elif command in ["good morning"]:
        response = "Good morning! Hope you have a great day."
        speak(response)
        return

    elif command in ["good evening"]:
        response = "Good evening! How was your day?"
        speak(response)
        return

    elif command in ["good night"]:
        response = "Good night! Sweet dreams."
        speak(response)
        return

    elif command in ["what is your name", "who are you"]:
        response = "My name is ISHA... INTELLIGENT SYSTEM FOR HUMAN ASSISTANCE"
        speak(response)
        return


### -------------------------------------------------------- ###

def open_folder_by_shortcut(command):
    folder_shortcuts = {
        "C12": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C13": r"C:\Users\acer\Desktop\cord\engine",
        "14": r"C:\Windows",
        "15": r"D:\\",  # Agar aapka D drive hai to
        "16": r"C:\\",  # Root C drive
    }

    command = command.lower().strip()
    
    if command.startswith("opening"):
        shortcut = command.replace("opening", "").strip()
        if shortcut in folder_shortcuts:
            path = folder_shortcuts[shortcut]
            speak(f"Opening folder for shortcut {shortcut}")
            os.startfile(path)
            return True
        else:
            speak("Shortcut number not found.")
            return False


def start_isha():
    get_installed_apps()
    wish_me()
    speak("I AM ISHA.... INTELLIGENT SYSTEM FOR HUMAN ASSISTANCE.")
    speak("Isha Assistant ready. Say Isha to activate.")
    while True:
        text = listen()
        if WAKE_WORD in text:
            speak("Yes, I am listening.")
            command = listen()
            if command:
                threading.Thread(target=handle_command, args=(command,)).start()

def start_thread():
    threading.Thread(target=start_isha).start()

def handle_entry_command(event=None):
    command = entry.get().strip()
    if command:
        log("You (typed): " + command)
        entry.delete(0, tk.END)
        threading.Thread(target=handle_command, args=(command,)).start()

start_button = tk.Button(window, text="🎙️ Voice", command=start_thread, bg="yellow")
start_button.pack(pady=10)

entry.bind("<FocusIn>", lambda e: entry.delete(0, tk.END) if entry.get() == "Type your command here" else None)
entry.bind("<Return>", handle_entry_command)
window.mainloop()
