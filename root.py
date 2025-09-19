import subprocess
import time

packages = [
    "tkintertable",
    "pyttsx3",
    "SpeechRecognition",
    "PyAudio",
    "PyAutoGUI",
    "selenium",
    "webdriver-manager",
    "pywhatkit",
    "requests",
    "sympy",
    "python-dotenv",
    "opencv-python",
    "numpy"
]

def install_package(pkg, retries=2):
    
    for attempt in range(retries + 1):
        print(f"\nInstalling {pkg}, Attempt {attempt + 1} ...")
        proc = subprocess.Popen(
            f'start cmd /k "pip install {pkg}"',
            shell=True
        )
        time.sleep(10)
        

for pkg in packages:
    install_package(pkg, retries=2)
    time.sleep(1)
