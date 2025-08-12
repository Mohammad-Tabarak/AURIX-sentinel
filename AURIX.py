import time
import speech_recognition as sr
import psutil
import webbrowser
import os
import subprocess
import ctypes
from googletrans import Translator
import requests
import re
import pyttsx3
import pyautogui

API_KEY = "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"
translator = Translator()

# ---------- Text Cleaning ----------
def clean_text(text):
    text = re.sub(r"\*\*|__|~~|`{1,3}", "", text)
    text = re.sub(r"\[.*?\]\(.*?\)", "", text)
    text = re.sub(r"\n+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# ---------- Text-to-Speech ----------
def speak(text, lang='en'):
    engine = pyttsx3.init('sapi5')
    clean = clean_text(text)

    if lang != 'en':
        try:
            translated = translator.translate(clean, dest=lang).text
        except Exception:
            translated = clean
    else:
        translated = clean

    if translated.strip():
        try:
            engine.say(translated)
            engine.runAndWait()
            time.sleep(0.3)
        except Exception as e:
            print(f"[ERROR] Speaking failed: {e}")

def speak_long_text(text, lang='en', max_chunk=300):
    text = clean_text(text).strip()
    while text:
        chunk = text[:max_chunk]
        last_period = chunk.rfind('.')
        if last_period != -1 and last_period > max_chunk // 2:
            chunk = chunk[:last_period+1]
        speak(chunk, lang)
        text = text[len(chunk):].strip()

# ---------- Speech Recognition ----------
def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎤 Listening...")
        r.pause_threshold = 0.8
        r.energy_threshold = 300
        try:
            audio = r.listen(source, timeout=5)
        except sr.WaitTimeoutError:
            print("[ERROR] Listening timed out")
            return ""
    try:
        query = r.recognize_google(audio)
        print(f"🗣️ You said: {query}")
        return query
    except Exception as e:
        print(f"[ERROR] Speech recognition failed: {e}")
        return ""

# ---------- AI Response ----------
def ask_deepseek(prompt, lang='en'):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "deepseek/deepseek-r1",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 500
    }
    try:
        response = requests.post(url, json=data, headers=headers)
        res_json = response.json()
        if "choices" in res_json and len(res_json["choices"]) > 0:
            reply = res_json['choices'][0]['message']['content']
            clean_reply = clean_text(reply)
            print(f"🤖 Jarvis (reply): {clean_reply}")
            speak_long_text(clean_reply, lang)
        else:
            speak("I encountered an error with the API.", lang)
    except Exception as e:
        print(f"[ERROR] API request failed: {e}")
        speak("Sorry, I could not connect to the API.", lang)

# ---------- PC Control ----------
def pc_control(command):
    cmd = command.lower()

    apps = {
        "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
    }
    kill_apps = {
        "chrome": "chrome.exe",
        "excel": "EXCEL.EXE",
        "notepad": "notepad.exe",
    }

    # Open apps dynamically
    for app_name in apps:
        if f"open {app_name}" in cmd:
            path = apps[app_name]
            try:
                subprocess.Popen(path)
                return f"Opening {app_name}"
            except Exception as e:
                return f"Failed to open {app_name}: {e}"

    # Close apps dynamically
    for app_name in kill_apps:
        if f"close {app_name}" in cmd:
            exe_name = kill_apps[app_name]
            try:
                os.system(f"taskkill /f /im {exe_name}")
                return f"Closing {app_name}"
            except Exception as e:
                return f"Failed to close {app_name}: {e}"

    # Volume control
    if "volume up" in cmd:
        pyautogui.press("volumeup")
        return "Volume increased"
    if "volume down" in cmd:
        pyautogui.press("volumedown")
        return "Volume decreased"
    if "mute" in cmd:
        pyautogui.press("volumemute")
        return "Volume muted"

    # Switch windows
    if "switch window" in cmd or "next window" in cmd:
        pyautogui.keyDown('alt')
        pyautogui.press('tab')
        pyautogui.keyUp('alt')
        return "Switched window"

    # Move mouse
    if "move mouse" in cmd:
        if "top left" in cmd:
            pyautogui.moveTo(0, 0)
            return "Moved mouse to top left"
        if "center" in cmd:
            width, height = pyautogui.size()
            pyautogui.moveTo(width / 2, height / 2)
            return "Moved mouse to center"

    # Shutdown, restart, lock
    if "shutdown" in cmd:
        speak("Shutting down the computer", "en")
        os.system("shutdown /s /t 5")
        return "Shutting down"
    if "restart" in cmd:
        speak("Restarting the computer", "en")
        os.system("shutdown /r /t 5")
        return "Restarting"
    if "lock" in cmd:
        speak("Locking the computer", "en")
        ctypes.windll.user32.LockWorkStation()
        return "Locking computer"

    return None

# ---------- Main task handler ----------
def perform_task(command, lang='en'):
    response = pc_control(command)
    if response:
        speak(response, lang)
        return

    command_lower = command.lower()
    if "open google" in command_lower:
        webbrowser.open("https://www.google.com")
        speak("Opening Google", lang)
    elif "open youtube" in command_lower:
        webbrowser.open("https://www.youtube.com")
        speak("Opening YouTube", lang)
    elif "search for" in command_lower:
        search_term = command_lower.replace("search for", "").strip()
        if search_term:
            webbrowser.open(f"https://www.google.com/search?q={search_term}")
            speak(f"Searching for {search_term}", lang)
    elif "battery" in command_lower:
        battery = psutil.sensors_battery()
        if battery:
            speak(f"Battery is at {battery.percent} percent", lang)
        else:
            speak("Battery information not available", lang)
    elif "time" in command_lower:
        current_time = time.strftime("%I:%M %p")
        speak(f"The time is {current_time}", lang)
    else:
        ask_deepseek(command, lang)

# ---------- Language Detection ----------
def detect_language(text):
    text = text.lower()
    if "hindi" in text:
        return 'hi'
    if "japanese" in text:
        return 'ja'
    if "russian" in text:
        return 'ru'
    if "chinese" in text:
        return 'zh-cn'
    if "spanish" in text:
        return 'es'
    return 'en'

# ---------- Main Loop ----------
def main():
    speak("Hello, I am AURIX. How can I help you?")
    while True:
        command = take_command()
        if not command:
            continue
        if command.lower() in ["exit", "stop", "quit"]:
            speak("Goodbye!", detect_language(command))
            break
        print(f"Command received: {command}")
        lang = detect_language(command)
        perform_task(command, lang)

if __name__ == "__main__":
    main()
