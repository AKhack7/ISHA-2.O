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
import datetime
import random

# ----------------------------- CONFIG ----------------------------- #
WAKE_WORD = "isha"
API_KEY = "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"
LANGUAGE_TRANSLATE_TO = "hi"
USE_ONLINE = True
SONG_FOLDER = os.path.expanduser("~/Music")

SONG_LINKS = [
    "https://youtu.be/s4KzN7mW8T8?si=1XeSDRT4Phi0uPc0",
    "https://youtu.be/9KDbPOg6hPE?si=ghw3R6h2Hw12wmKk",
    "https://youtu.be/JgDNFQ2RaLQ?si=qkYjWJPMPYvnmYlN",
    "https://youtu.be/n2dVFdqMYGA?si=h9ZOTspiDKpiw6oW",
    "https://youtu.be/sUf2PtEZris?si=HA5DmCB0dX4wg6HO",
    "https://youtu.be/8of5w7RgcTc?si=-XXx0v8gXJmABZ57",
    "https://youtu.be/aHuuaIAS_U4?si=hsg_65_lfv316AmK",
    "https://youtu.be/TkAiQJzctFY?si=uBV5IdHyfHqbBOWr",
    "https://youtu.be/UyoDdroSXXs?si=lIyPi9uJzF1DgMi7",
    "https://youtu.be/-2RAq5o5pwc?si=Cp6V5C4r9qKFz4QX",
    "https://youtu.be/43KpwU2tQzM?si=cz3D_rRqk5gGaDmk",
    "https://youtu.be/YJg1rs0R2sE?si=aXMKCjm0Qv965_BC",
    "https://youtu.be/cY2oEfF9ggM?si=T2Y3BH7sJuUix-93"
    # ... add more as needed
]

# ----------------------------- OPENROUTER SETUP ----------------------------- #
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=API_KEY,
)

# ----------------------------- SETUP ----------------------------- #
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 175)

shortcut_states = {}
recognizer = sr.Recognizer()
installed_apps = set()

# ----------------------------- GUI SETUP ----------------------------- #
window = tk.Tk()
window.title("ISHA")
window.geometry("600x500")
window.configure(bg="#ADD8E6")

log_box = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=70, height=20, font=("Arial", 10, "italic"), fg="black", bg="#FAFAFA", border=3, relief="solid")
log_box.pack(pady=10)

status_label = tk.Label(window, text="Status: Ready", fg="green", bg="#ADD8E6")
status_label.pack()

entry = tk.Entry(window, width=50)
entry.pack(pady=5)
entry.insert(0, "Type your command here")

entry.bind("<FocusIn>", lambda e: entry.delete(0, tk.END) if entry.get() == "Type your command here" else None)

# ----------------------------- FUNCTIONS ----------------------------- #
def log(message):
    log_box.insert(tk.END, message + "\n")
    log_box.see(tk.END)

def speak(text):
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
        "word": "winword",
        "cmd":"command Prompt"
    }
    website_map = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com",
        "facebook": "https://www.facebook.com",
        "gmail": "https://mail.google.com",
        "whatsapp": "https://web.whatsapp.com"
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

def play_songs():
    if is_connected():
        url = random.choice(SONG_LINKS)
        speak("Playing a random song on YouTube")
        webbrowser.open(url)
    else:
        try:
            files = [f for f in os.listdir(SONG_FOLDER) if f.endswith(('.mp3', '.wav'))]
            if files:
                song_path = os.path.join(SONG_FOLDER, files[0])
                os.startfile(song_path)
                speak("Playing song from your computer")
            else:
                speak("No songs found in your Music folder")
        except:
            speak("Could not play song")

def online_answer(prompt):
    try:
        completion = client.chat.completions.create(
            extra_headers={
                "HTTP-Referer": "http://localhost",
                "X-Title": "Isha AI Assistant",
            },
            model="openai/gpt-4o",
            max_tokens=100,
            messages=[{"role": "user", "content": prompt}]
        )
        return completion.choices[0].message.content.strip()
    except Exception as e:
        return f"I'm having trouble reaching OpenRouter. Error: {str(e)}"

def translate_speak(text, target_lang=LANGUAGE_TRANSLATE_TO):
    try:
        translated = GoogleTranslator(source='auto', target=target_lang).translate(text)
        speak(translated)
    except:
        speak("Translation failed.")

def handle_command(command):
    command = command.lower()
    if any(x in command for x in ["time", "samy"]):
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        response = f"Current time is {current_time}"
    elif any(x in command for x in ["date", "tarik"]):
        today = datetime.datetime.now().strftime("%d %B %Y")
        response = f"Today's date is {today}"
    elif command.strip() in ["hi", "hello", "hey"]:
        response = "Hello Milan! How can I assist you today?"
    elif command.strip() in ["good morning"]:
        response = "Good morning Milan! Hope you have a great day."
    elif command.strip() in ["good evening"]:
        response = "Good evening Milan! How was your day?"
    elif command.strip() in ["good night"]:
        response = "Good night Milan! Sweet dreams."
    elif command.strip() in ["what is your name", "who are you"]:
        response = "My name is ISHA... INTELLIGENT SYSTEM FOR HUMAN ASSISTANCE"
    elif any(op in command for op in ["+", "-", "*", "/"]):
        try:
            result = eval(command)
            response = f"The answer is {result}"
        except:
            response = "I couldn't calculate that."
    elif command.startswith("open"):
        app_name = command.replace("open", "").strip()
        launch_app(app_name)
        return
    elif "live translation" in command:
        speak("Live translation started. Speak now.")
        live_translate_mode()
        return
    elif "play song" in command or "play music" in command:
        play_songs()
        return
    elif is_connected() and USE_ONLINE:
        response = online_answer(command)
    else:
        response = "I can't answer that offline."
    log("Isha: " + response)
    speak(response)

def live_translate_mode():
    while True:
        text = listen()
        if any(x in text for x in ["exit", "stop"]):
            speak("Translation ended.")
            break
        translate_speak(text)

def start_isha():
    get_installed_apps()
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

# ----------------------------- RUN GUI ----------------------------- #
start_button = tk.Button(window, text="Voice", command=start_thread, bg="yellow")
start_button.pack(pady=10)

entry.bind("<Return>", handle_entry_command)
window.mainloop()
