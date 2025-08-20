# [IMPORTS]
import speech_recognition as sr
import webbrowser
import pyttsx3
import MusicLibrary
import datetime
import os
import psutil
import g4f
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import pyautogui
import ctypes
import time
import pywhatkit

# [INITIALIZE]
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# [CORE SPEECH METHODS]
def takeCommand():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        print("Listening...")
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=7)
        try:
            command = recognizer.recognize_google(audio)
            print(f"Command recognized: {command}")
            return command.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
        return ""

def speak(text):
    engine.say(text)
    engine.runAndWait()

# [HELPERS]
def calculate_expression(command):
    command = command.lower().replace("what is", "").replace("calculate", "").strip()
    command = command.replace("plus", "+").replace("minus", "-").replace("times", "*")
    command = command.replace("multiplied by", "*").replace("x", "*")
    command = command.replace("divided by", "/").replace("over", "/")

    try:
        result = eval(command)
        speak(f"The answer is {result}")
    except Exception:
        speak("Sorry, I couldn't calculate that.")

def takeScreenshot():
    screenshot = pyautogui.screenshot()
    screenshot.save("screenshot.png")
    speak("Screenshot taken and saved as screenshot.png")

def volumeControl(action):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    if action == "increase":
        current_volume = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(min(current_volume + 0.1, 1.0), None)
        speak("Volume increased.")
    elif action == "decrease":
        current_volume = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(max(current_volume - 0.1, 0.0), None)
        speak("Volume decreased.")
    elif action == "mute":
        volume.SetMute(1, None)
        speak("Volume muted.")

def close_browser():
    browsers = ["chrome.exe", "firefox.exe", "msedge.exe", "brave.exe"]
    closed = False
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() in browsers:
            proc.kill()
            closed = True
    speak("Browser closed." if closed else "No browser was open.")

# [COMMAND ROUTING]
def handle_greetings(c):
    if c == "jarvis":
        speak("Yes sir. I am here.")

def handle_music_commands(c):
    if c.lower().startswith("play"):
        song = c.lower().replace("play", "", 1).strip()
        link = MusicLibrary.music.get(song)
        if link:
            webbrowser.open(link)
        else:
            speak(f"{song} not found in the library.")

def handle_web_commands(c):
    if "open google" in c.lower():
        speak("Opening Google")
        webbrowser.open("https://www.google.com")
    elif "open youtube" in c:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")
    elif "search" in c:
        query = c.replace("search", "")
        url = f"https://www.google.com/search?q={query}"
        webbrowser.open(url)

def handle_datetime_commands(c):
    if "time" in c:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"Sir, the time is {strTime}")
    elif "date" in c:
        strDate = datetime.datetime.now().strftime("%Y-%m-%d")
        speak(f"Sir, today's date is {strDate}")

def handle_calculator(c):
    if "calculate" in c or "what is" in c:
        speak("Sure, I can help with that. What would you like to calculate?")
        c = takeCommand()
        calculate_expression(c)

def handle_volume_commands(c):
    if "volume up" in c:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevel(volume.GetMasterVolumeLevel() + 1.0, None)
    elif "volume down" in c:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevel(volume.GetMasterVolumeLevel() - 1.0, None)
    elif "mute" in c:
        pyautogui.press("volumemute")

def handle_screen_and_typing(c):
    if "take screenshot" in c:
        screenshot = pyautogui.screenshot()
        screenshot.save("screenshot.png")
        speak("Screenshot taken and saved as screenshot.png")
    elif "type" in c:
        text = c.replace("type", "")
        pyautogui.write(text)
    elif "scroll up" in c:
        pyautogui.scroll(300)
    elif "scroll down" in c:
        pyautogui.scroll(-300)

def handle_system_commands(c):
    if "shutdown" in c:
        speak("Shutting down the system")
        os.system("shutdown /s /t 1")
    elif "restart" in c:
        speak("Restarting the system")
        os.system("shutdown /r /t 1")
    elif "log off" in c:
        speak("Logging off")
        os.system("shutdown -l")

def handle_app_commands(c):
    if "open notepad" in c:
        speak("Opening Notepad")
        os.system("notepad")
    elif "open command prompt" in c:
        speak("Opening Command Prompt")
        os.system("start cmd")
    elif "open camera" in c:
        speak("Opening Camera")
        os.system("start microsoft.windows.camera:")

def handle_chatgpt_query(c):
    if "question" in c:
        question = c.replace("question", "")
        response = g4f.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": question}]
        )
        print(response)
        speak(response)

def processCommand():
    c = takeCommand()
    handle_greetings(c)
    handle_music_commands(c)
    handle_web_commands(c)
    handle_datetime_commands(c)
    handle_calculator(c)
    handle_volume_commands(c)
    handle_screen_and_typing(c)
    handle_system_commands(c)
    handle_app_commands(c)
    handle_chatgpt_query(c)
    speak("I didn't understand that command.")

# [MAIN LOOP]
if __name__ == "__main__":
    speak("Initializing Jarvis...")

    while True:
        print("Listening for wake word...")
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=1)
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=3)
                command = recognizer.recognize_google(audio)
                print(f"Heard: {command}")
        except sr.WaitTimeoutError:
            continue
        except sr.UnknownValueError:
            continue
        except sr.RequestError as e:
            print(f"API error: {e}")
            continue

        if command.lower() == "jarvis":
            speak("Yup! I'm listening. Say 'stop' to exit.")
            while True:
                try:
                    with sr.Microphone() as source:
                        recognizer.adjust_for_ambient_noise(source, duration=1)
                        print("Listening for command...")
                        audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
                        command = recognizer.recognize_google(audio).lower()
                        print(f"Command: {command}")
                        if command in ["stop", "exit", "bye"]:
                            speak("Okay, going back to sleep.")
                            break
                        processCommand()
                except sr.UnknownValueError:
                    speak("Sorry, I didn't catch that.")
                except sr.RequestError as e:
                    speak(f"Error from recognition service: {e}")
                except sr.WaitTimeoutError:
                    print("Timeout waiting for speech.")
