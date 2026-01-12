try:
    import speech_recognition as sr
except Exception:
    sr = None
try:
    import pyaudio  
except Exception:
    pyaudio = None
import os
import webbrowser
from config import apikey, OPENROUTER_BASE_URL, OPENROUTER_MODEL
import datetime
import random
import re
import sys
import requests
import string
import subprocess
import math

try:
    from PIL import ImageGrab
except Exception:
    ImageGrab = None

# Windows API available check (pywin32 package)
HAS_WIN32 = os.name == "nt"


chatStr = ""
# Windows app launcher helpers
def _win_try_start(command: str) -> bool:
    try:
        exit_code = os.system(command)
        return exit_code == 0
    except Exception:
        return False

def _win_try_paths(possible_paths):
    for p in possible_paths:
        # Expand %USERNAME% style and env vars
        expanded = os.path.expandvars(p)
        if os.path.exists(expanded):
            try:
                os.startfile(expanded)
                return True
            except Exception:
                pass
    return False

def open_app_windows(query: str) -> bool:
    q = query.lower()
    if "open chrome" in q:
        if _win_try_start("start chrome"):
            return True
        return _win_try_paths([
            r"C:%HOMEPATH%\AppData\Local\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ])
    if "open opera" in q or "open opera browser" in q:
        if _win_try_start("start opera"): 
            return True
        return _win_try_paths([
            r"C:%HOMEPATH%\AppData\Local\Programs\Opera\launcher.exe",
            r"C:\Program Files\Opera\launcher.exe",
            r"C:\Program Files (x86)\Opera\launcher.exe"
        ])
    if "open edge" in q or "open microsoft edge" in q:
        if _win_try_start("start msedge"):
            return True
        return _win_try_paths([
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ])
    if "open firefox" in q:
        if _win_try_start("start firefox"):
            return True
        return _win_try_paths([
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        ])
    if "open whatsapp" in q:
        if _win_try_start("start whatsapp:"):
            return True
        return _win_try_paths([
            r"%LOCALAPPDATA%\WhatsApp\WhatsApp.exe",
            r"%USERPROFILE%\AppData\Local\WhatsApp\WhatsApp.exe"
        ])
    if "open vscode" in q or "open vs code" in q or "open code" in q:
        if _win_try_start("start code"):
            return True
        return _win_try_paths([
            r"%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe"
        ])
    if "open spotify" in q:
        if _win_try_start("start spotify:"):
            return True
        return _win_try_paths([
            r"%APPDATA%\Spotify\Spotify.exe"
        ])
    return False

# Calculator function
def calculate_math(query):
    try:
        # Extract math expression from query
        q = query.lower()
        if "calculate" in q or "compute" in q or "what is" in q:
            # Replace text numbers and operations
            expr = q.replace("calculate", "").replace("compute", "").replace("what is", "").strip()
            expr = expr.replace("times", "*").replace("multiply", "*").replace("multiplied by", "*")
            expr = expr.replace("plus", "+").replace("add", "+").replace("added to", "+")
            expr = expr.replace("minus", "-").replace("subtract", "-").replace("subtracted from", "-")
            expr = expr.replace("divided by", "/").replace("divide", "/").replace("over", "/")
            expr = expr.replace("power", "**").replace("to the power of", "**")
            expr = expr.replace("squared", "**2").replace("cubed", "**3")
            
            # Remove non-math characters for safety
            safe_expr = re.sub(r'[^0-9+\-*/().\s]', '', expr)
            if safe_expr:
                result = eval(safe_expr)
                return str(result)
    except Exception:
        pass
    return None

# Screenshot function
def take_screenshot():
    try:
        if ImageGrab:
            img = ImageGrab.grab()
            filename = f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            img.save(filename)
            return filename
        else:
            # Fallback using Windows Snipping Tool
            if os.name == "nt":
                os.system("snippingtool /clip")
                return "Screenshot copied to clipboard"
    except Exception as e:
        return f"Error: {e}"
    return None

# Volume control feature removed per user request

# Close/Minimize window function
def control_windows(query):
    try:
        q = query.lower()
        if os.name == "nt":
            if "close chrome" in q:
                # More reliable taskkill
                os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
                return "Chrome closed"
            elif "close opera" in q:
                os.system("taskkill /F /IM opera.exe /T >nul 2>&1")
                return "Opera closed"
            elif "close firefox" in q:
                os.system("taskkill /F /IM firefox.exe /T >nul 2>&1")
                return "Firefox closed"
            elif "close edge" in q:
                os.system("taskkill /F /IM msedge.exe /T >nul 2>&1")
                return "Edge closed"
            elif "minimize window" in q or "minimize" in q:
                # Alternative minimize command
                os.system("powershell -Command \"$wsh = New-Object -ComObject WScript.Shell; $wsh.SendKeys('% ')\"")
                return "Window minimized"
    except Exception as e:
        return f"Window control error: {e}"
    return None

# System control function (NEW)
def control_system(query):
    try:
        q = query.lower()
        if os.name == "nt":
            if "shutdown" in q:
                # Shutdown with delay
                os.system("shutdown /s /t 5")
                return "Shutting down in 5 seconds"
            elif "restart" in q or "reboot" in q:
                os.system("shutdown /r /t 5")
                return "Restarting in 5 seconds"
            elif "sleep" in q or "hibernate" in q:
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                return "Going to sleep"
            elif "lock" in q or "lock screen" in q:
                os.system("rundll32.exe user32.dll,LockWorkStation")
                return "Screen locked"
    except Exception as e:
        return f"System control error: {e}"
    return None

# Dictionary/Word meaning function
def get_word_meaning(query):
    try:
        q = query.lower()
        if "meaning of" in q or "meaning" in q or "define" in q:
            word = re.sub(r"(meaning of|meaning|define)", "", q).strip()
            if word:
                # Using Free Dictionary API
                url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}"
                resp = requests.get(url, timeout=10)
                if resp.status_code == 200:
                    data = resp.json()[0]
                    meanings = data.get("meanings", [])
                    if meanings:
                        definition = meanings[0]["definitions"][0]["definition"]
                        return f"{word}: {definition}"
                    return f"Meaning not found for {word}"
                else:
                    # Fallback: Use AI to explain
                    return None
    except Exception:
        pass
    return None


# Web search function
def web_search(query):
    try:
        q = query.lower()
        if "search" in q:
            search_term = q.replace("search", "").strip()
            if search_term:
                url = f"https://www.google.com/search?q={search_term.replace(' ', '+')}"
                webbrowser.open(url)
                return f"Searching for {search_term}"
    except Exception:
        pass
    return None

#  ai chat system
def chat(query):
    global chatStr
    print(chatStr)
    chatStr += f"User: {query}\nNoha: "
    try:
        payload = {
            "model": OPENROUTER_MODEL,
            "messages": [
                {"role": "system", "content": "You are Noha, a helpful voice assistant."},  
                {"role": "user", "content": chatStr}
            ],
            "temperature": 0.7,
            "max_tokens": 256,
            "top_p": 1
        }
        headers = {
            "Authorization": f"Bearer {apikey}",
            "Content-Type": "application/json"
        }
        resp = requests.post(f"{OPENROUTER_BASE_URL}/chat/completions", json=payload, headers=headers, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        answer = data["choices"][0]["message"]["content"]
        say(answer)
        chatStr += f"{answer}\n"
        return answer
    except Exception as e:
        err = f"Sorry, an error occurred: {e}"
        say(err)
        return err

#ai file generator
def ai(prompt):
    text = f"OpenRouter (DeepSeek) response for Prompt: {prompt} \n *************************\n\n"
    try:
        payload = {
            "model": OPENROUTER_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 256,
            "top_p": 1
        }
        headers = {
            "Authorization": f"Bearer {apikey}",
            "Content-Type": "application/json"
        }
        resp = requests.post(f"{OPENROUTER_BASE_URL}/chat/completions", json=payload, headers=headers, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        answer = data["choices"][0]["message"]["content"]
        text += answer
    except Exception as e:
        text += f"Error: {e}"

    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    # Safe filename from prompt
    base = re.sub(r"[^a-zA-Z0-9_-]+", " ", prompt).strip()
    if not base:
        base = f"prompt-{random.randint(1, 2343434356)}"
    else:
        base = "-".join(base.split())[:60]

    with open(f"Openai/{base}.txt", "w", encoding="utf-8") as f:
        f.write(text)

def say(text):
    try:
        # Try pyttsx3 first
        import pyttsx3
        engine = pyttsx3.init()
        engine.setProperty('rate', 150)  # Speed of speech
        engine.setProperty('volume', 0.9)  # Volume level
        engine.say(text)
        engine.runAndWait()
    except Exception as e1:
        try:
            # Try edge-tts for Windows 10/11
            if os.name == "nt":
                import asyncio
                import edge_tts
                
                async def _speak():
                    tts = edge_tts.Communicate(text=text, voice="en-US-AriaNeural")
                    await tts.save("temp_speech.mp3")
                    os.system("start temp_speech.mp3")
                
                asyncio.run(_speak())
            else:
                # macOS 'say' command
                os.system(f'say "{text}"')
        except Exception as e2:
            # Final fallback to console output
            print(f"Noha: {text}")

# voice command 
def takeCommand():
    if sr is None:
        try:
            return input("Type your command: ")
        except Exception:
            return ""
    # If PyAudio is not installed/available, fall back to text input immediately
    if pyaudio is None:
        try:
            return input("Mic not available. Type your command: ")
        except Exception:
            return ""
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.5)
            r.pause_threshold = 1
            print("Listening...")
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "Could not understand audio"
        except sr.RequestError as e:
            return f"Could not request results; {e}"
        except Exception as e:
            return f"Error: {e}"
    except (AttributeError, OSError):
        # PyAudio/microphone not available: fall back to text input
        try:
            return input("Mic not available. Type your command: ")
        except Exception:
            return ""
    except Exception as e:
        return f"Listening error: {e}"
    
# main loop and command handling

if __name__ == '__main__':
    print('Welcome to Noha A.I')  # Changed from Jarvis to Noha
    say("Hello, I am Noha. How can I help you today?")  # Changed greeting
    while True:
        print("Listening...")
        query = takeCommand()

        # Normalize and sanitize recognized text
        if not isinstance(query, str):
            continue
        q = query.strip().lower()

        # Ignore recognizer status or error messages
        if not q:
            continue
        ignore_starts = ("could not", "listening error", "mic not available", "error:")
        if q == "could not understand audio" or q.startswith(ignore_starts):
            continue

        # Check for exit command
        if q in ["exit", "quit", "goodbye", "bye", "stop"]:
            say("Goodbye! Have a great day!")
            break
        elif "noha quit" in q or "exit noha" in q:
            say("Goodbye!")
            break

        # todo: Add more sites
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],]
        _opened_site = False
        for site in sites:
            if f"open {site[0]}" in q:
                say(f"Opening {site[0]}...")
                webbrowser.open(site[1])
                _opened_site = True
                break
        if _opened_site:
            continue

        # Windows desktop app commands
        if os.name == "nt" and "open notepad" in q:
            os.system("start notepad")
            say("Opening Notepad")
            continue
        elif os.name == "nt" and ("open calculator" in q or "open calc" in q):
            os.system("start calc")
            say("Opening Calculator")
            continue
        elif os.name == "nt" and ("open cmd" in q or "open command prompt" in q):
            os.system("start cmd")
            say("Opening Command Prompt")
            continue
        elif os.name == "nt" and "open camera" in q:
            os.system("start microsoft.windows.camera:")
            say("Opening Camera")
            continue
        elif os.name == "nt":
            # Try popular apps via helper
            if open_app_windows(q):
                say("Opening application")
                continue

        # Music player
        if "play music" in q or "open music" in q:
            musicPath = "C:\\Users\\Public\\Music\\Sample Music"  # Update this path
            try:
                if os.name == "nt":
                    os.system(f'explorer "{musicPath}"')
                    say("Opening music folder")
                else:
                    os.system(f"open {musicPath}")
            except Exception:
                say("Music path not found")
            continue

        elif "the time" in q:
            hour = datetime.datetime.now().strftime("%H")
            minute = datetime.datetime.now().strftime("%M")
            say(f"The time is {hour} hours and {minute} minutes")
            print(f"Time: {hour}:{minute}")
            continue

        elif "using artificial intelligence" in q:
            ai(prompt=query)
            continue

        elif "reset chat" in q:
            chatStr = ""
            say("Chat history cleared")
            continue

        # Calculator
        elif "calculate" in q or "compute" in q or ("what is" in q and any(op in q for op in ["times", "plus", "minus", "divided", "multiply"])):
            result = calculate_math(query)
            if result:
                say(f"The answer is {result}")
                print(f"Result: {result}")
            continue

        # Screenshot
        elif "take screenshot" in q or "screenshot" in q:
            result = take_screenshot()
            if result:
                say(f"Screenshot saved")
                print(f"Screenshot: {result}")
            continue

        # NEW: System control (shutdown, restart, etc.)
        elif "shutdown" in q:
            # confirmation
            say("Do you want to shut down the computer now? Say yes to confirm.")
            conf = takeCommand()
            if isinstance(conf, str):
                conf_norm = conf.strip().lower()
                print(f"Confirmation heard: {conf_norm}")
                affirmatives = ("y","yes","ya","yah","haan","ha","ok","okay","sure","theek","thik","bilkul")
                if any(conf_norm.startswith(a) or a in conf_norm for a in affirmatives):
                    result = control_system(q)
                    if result:
                        say(result)
                        print(result)
                else:
                    say("Shutdown cancelled")
            else:
                say("Shutdown cancelled")
            continue
        elif "restart" in q or "reboot" in q:
            say("Do you want to restart the computer now? Say yes to confirm.")
            conf = takeCommand()
            if isinstance(conf, str):
                conf_norm = conf.strip().lower()
                print(f"Confirmation heard: {conf_norm}")
                affirmatives = ("y","yes","ya","yah","haan","ha","ok","okay","sure","theek","thik","bilkul")
                if any(conf_norm.startswith(a) or a in conf_norm for a in affirmatives):
                    result = control_system(q)
                    if result:
                        say(result)
                        print(result)
                else:
                    say("Restart cancelled")
            else:
                say("Restart cancelled")
            continue
        elif "sleep" in q or "hibernate" in q:
            result = control_system(q)
            if result:
                say(result)
                print(result)
            continue
        elif "lock" in q or "lock screen" in q:
            result = control_system(q)
            if result:
                say(result)
                print(result)
            continue

        # Window/app control
        elif "close chrome" in q or "close opera" in q or "close firefox" in q or "close edge" in q:
            result = control_windows(q)
            if result:
                say(result)
                print(result)
            continue
        elif "minimize window" in q or "minimize" in q:
            result = control_windows(q)
            if result:
                say(result)
                print(result)
            continue

        # Dictionary/Word meaning
        elif "meaning of" in q or ("meaning" in q and len(query.split()) <= 5) or "define" in q:
            result = get_word_meaning(query)
            if result:
                say(result)
                print(result)
            continue

        # Web search
        elif "search" in q:
            result = web_search(query)
            if result:
                say(result)
                print(result)
            continue

        # Default: Chat mode
        else:
            print("Chatting...")
            chat(query)