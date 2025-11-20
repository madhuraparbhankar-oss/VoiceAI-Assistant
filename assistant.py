import os
import time
import speech_recognition as sr
import pyttsx3
from time import sleep
import pyautogui
from datetime import datetime
import threading
import webbrowser
import screen_brightness_control as sbc
import wikipedia
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import pygetwindow as gw
import psutil
import requests
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import pyperclip
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import comtypes
import pyjokes
from google import genai


# =========================
# GEMINI API CONFIGURATION
# =========================

GEMINI_API_KEY = "add your gemini api key"  # Your API key

try:
    # New Gemini client (no configure())
    client = genai.Client(api_key=GEMINI_API_KEY)
    GEMINI_ENABLED = True
    print("✨ Gemini API initialized successfully")
except Exception as e:
    print(f"Gemini API initialization failed: {e}")
    GEMINI_ENABLED = False


# =========================
# GEMINI RESPONSE FUNCTION
# =========================
def ask_gemini(question):
    """Ask Gemini AI and get a response using new API"""
    if not GEMINI_ENABLED:
        return "Gemini AI is not configured. Please add your API key."

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",   # New correct model
            contents=question
        )
        return response.text
    except Exception as e:
        return f"Gemini Error: {str(e)}"
    
    
# Initialize recognizer
recognizer = sr.Recognizer()
wake_word = "assistant"

# Function to handle text-to-speech
def speak(text):
    def run():
        try:
            comtypes.CoInitialize()  # Initialize COM in this thread
            engine = pyttsx3.init()
            voices = engine.getProperty('voices')
            for voice in voices:
                if "zira" in voice.name.lower():
                    engine.setProperty('voice', voice.id)
                    break
            engine.say(text)
            engine.runAndWait()
        except Exception as e:
            print(f"Speech engine error: {e}")
        finally:
            comtypes.CoUninitialize()  # Cleanup COM after use

    threading.Thread(target=run, daemon=True).start()
    



def set_volume(level):
    try:
        comtypes.CoInitialize()
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        precise_level = level / 100.0
        volume.SetMasterVolumeLevelScalar(precise_level, None)
        print(f"Volume successfully set to: {level}%")
        speak(f"Volume set to {level} percent")
    except Exception as e:
        print(f"Failed to set volume: {e}")
        speak("Failed to adjust volume")
    finally:
        try:
            comtypes.CoUninitialize()
        except:
            pass


def create_file(file_name):
    try:
        desktop_directory = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop_directory, file_name)
        with open(file_path, 'w') as file:
            file.write("")
        speak(f"File {file_name} created successfully on your Desktop.")
    except Exception as e:
        speak(f"Failed to create file: {e}")


def create_folder(folder_name):
    try:
        desktop_directory = os.path.join(os.path.expanduser("~"), "Desktop")
        folder_path = os.path.join(desktop_directory, folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            speak(f"Folder '{folder_name}' created successfully on your Desktop.")
        else:
            speak(f"Folder '{folder_name}' already exists on your Desktop.")
    except Exception as e:
        speak(f"Failed to create folder: {e}")


def rename_folder_on_desktop(old_folder_name, new_folder_name):
    try:
        desktop_directory = os.path.join(os.path.expanduser("~"), "Desktop")
        old_folder_path = os.path.join(desktop_directory, old_folder_name)
        new_folder_path = os.path.join(desktop_directory, new_folder_name)
        if os.path.exists(old_folder_path):
            os.rename(old_folder_path, new_folder_path)
            speak(f"Folder '{old_folder_name}' has been renamed to '{new_folder_name}' on your Desktop.")
        else:
            speak(f"Folder '{old_folder_name}' does not exist on your Desktop.")
    except Exception as e:
        speak(f"Failed to rename the folder: {e}")


def rename_file_on_desktop(old_name, new_name):
    try:
        desktop_directory = os.path.join(os.path.expanduser("~"), "Desktop")
        old_path = os.path.join(desktop_directory, old_name)
        new_path = os.path.join(desktop_directory, new_name)
        if os.path.exists(old_path):
            os.rename(old_path, new_path)
            speak(f"File '{old_name}' has been renamed to '{new_name}' on your Desktop.")
        else:
            speak(f"File '{old_name}' does not exist on your Desktop.")
    except Exception as e:
        speak(f"Failed to rename the file: {e}")


def set_brightness(level):
    try:
        if 0 <= level <= 100:
            sbc.set_brightness(level)
            speak(f"Brightness set to {level} percent")
            print(f"Brightness successfully set to: {level}%")
        else:
            speak("Please provide a brightness level between 0 and 100.")
    except Exception as e:
        speak(f"Failed to adjust brightness: {e}")
        print(f"Error: {e}")


# Function to listen for commands including wake word
listening = False

def listen():
    global listening
    listening = True
    
    try:
        with sr.Microphone() as source:
            print("🎤 Microphone initialized successfully")
            recognizer.adjust_for_ambient_noise(source, duration=1)
            update_status("🟢 Listening...", "#00ff88")
            log_message("✅ Microphone ready. Say 'assistant' followed by your command.")

            while listening:
                try:
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
                    command = recognizer.recognize_google(audio).lower()
                    
                    if wake_word in command:
                        command = command.replace(wake_word, "").strip()
                        log_message(f"🗣️ You: {command}")
                        execute_command(command)
                        
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    continue
                except sr.RequestError as e:
                    speak("Sorry, I couldn't connect to the speech recognition service.")
                    log_message(f"❌ Error: Could not connect to service")
                    break
                except Exception as e:
                    continue
                    
    except OSError as e:
        error_msg = "❌ Microphone Error: Please check microphone connection and permissions"
        log_message(error_msg)
        update_status("❌ Microphone Error", "#ff4444")
        speak("Microphone not found. Please check your microphone connection.")
    except Exception as e:
        error_msg = f"❌ Error initializing microphone: {str(e)}"
        log_message(error_msg)
        update_status("❌ Error", "#ff4444")


def stop_listening():
    global listening
    listening = False
    speak("Stopped listening.")
    log_message("⏹ Stopped listening.")
    update_status("⏹ Stopped", "#ff9800")


def execute_command(command):
    """Execute commands with AI fallback"""
    
    # Check for system commands first
    if "create file" in command:
        file_name = command.replace("create file", "").strip()
        create_file(file_name)

    elif "create folder" in command:
        folder_name = command.replace("create folder", "").strip()
        create_folder(folder_name)

    elif 'open' in command:
        exe_name = command.replace('open', '').strip()
        if not exe_name:
            speak("Please specify the application you want to open.")
        else:
            try:
                speak(f"Opening {exe_name}.")
                pyautogui.press('win')
                sleep(1)
                pyautogui.write(exe_name)
                sleep(1)
                pyautogui.press('enter')
                sleep(2)
            except Exception as e:
                speak(f"Failed to open {exe_name}.")

    elif "increase volume" in command:
        pyautogui.press("volumeup")
        speak("Increasing volume.")

    elif "decrease volume" in command:
        pyautogui.press("volumedown")
        speak("Decreasing volume.")

    elif 'shutdown' in command:
        speak("Shutting down the computer.")
        os.system('shutdown /s /t 1')

    elif "rename file" in command:
        parts = command.replace("rename file", "").strip().split(" to ")
        if len(parts) == 2:
            rename_file_on_desktop(parts[0].strip(), parts[1].strip())
        else:
            speak("Please specify both the old and new file names.")

    elif "rename folder" in command:
        parts = command.replace("rename folder", "").strip().split(" to ")
        if len(parts) == 2:
            rename_folder_on_desktop(parts[0].strip(), parts[1].strip())
        else:
            speak("Please specify both the old and new folder names.")

    elif "date" in command or "today" in command:
        current_date = datetime.now().strftime("%A, %B %d, %Y")
        speak(f"Today's date is {current_date}.")
        log_message(f"📅 {current_date}")

    elif "time" in command:
        current_time = datetime.now().strftime("%I:%M %p")
        speak(f"The time is {current_time}.")
        log_message(f"⏰ {current_time}")

    elif "set volume to" in command:
        try:
            words = command.split()
            if "to" in words:
                index = words.index("to") + 1
                level = int(words[index].replace("%", ""))
                if 0 <= level <= 100:
                    set_volume(level)
                else:
                    speak("Please provide a volume level between 0 and 100.")
        except (ValueError, IndexError):
            speak("Invalid volume level.")

    elif "set brightness to" in command:
        try:
            words = command.split()
            if "to" in words:
                index = words.index("to") + 1
                level = int(words[index].replace("%", ""))
                if 0 <= level <= 100:
                    set_brightness(level)
                else:
                    speak("Please provide a brightness level between 0 and 100.")
        except (ValueError, IndexError):
            speak("Invalid brightness level.")

    elif command.startswith("type "):
        text_to_type = command.replace("type ", "").strip()
        if text_to_type:
            pyautogui.write(text_to_type, interval=0.05)
            speak("Typing complete.")
        else:
            speak("I didn't catch what you wanted to type.")

    elif "search for" in command or "google" in command:
        query = command.replace("search for", "").replace("google", "").strip()
        if query:
            speak(f"Searching for {query} on Google.")
            webbrowser.open(f"https://www.google.com/search?q={query}")

    elif "youtube" in command:
        speak("Opening YouTube.")
        webbrowser.open("https://www.youtube.com")

    elif 'minimize window' in command or 'minimise window' in command:
        pyautogui.hotkey('win', 'down')
        speak("Minimizing window")

    elif 'maximize window' in command:
        pyautogui.hotkey('win', 'up')
        speak("Maximizing window")

    elif 'scroll up' in command:
        pyautogui.scroll(300)
        speak("Scrolling up")

    elif 'scroll down' in command:
        pyautogui.scroll(-300)
        speak("Scrolling down")

    elif "screenshot" in command:
        try:
            save_path = os.path.expanduser("~/Pictures/Screenshots")
            os.makedirs(save_path, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            screenshot_path = os.path.join(save_path, f'screenshot_{timestamp}.png')
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            speak("Screenshot saved")
            log_message(f"📸 Screenshot saved")
        except Exception as e:
            speak("Failed to take screenshot.")

    elif 'lock computer' in command:
        os.system('rundll32.exe user32.dll,LockWorkStation')
        speak("Locking the computer")

    elif "wikipedia" in command:
        query = command.replace("wikipedia", "").strip()
        speak(f"Searching Wikipedia for {query}.")
        try:
            result = wikipedia.summary(query, sentences=2)
            speak(result)
            log_message(f"📚 Wikipedia: {result[:100]}...")
        except Exception:
            speak("I couldn't find anything on Wikipedia.")

    elif "battery" in command:
        battery = psutil.sensors_battery()
        percentage = battery.percent
        speak(f"Battery is at {percentage} percent.")
        log_message(f"🔋 Battery: {percentage}%")

    elif "internet" in command:
        try:
            requests.get("http://www.google.com", timeout=3)
            speak("Internet is connected.")
            log_message("🌐 Internet: Connected")
        except requests.ConnectionError:
            speak("No internet connection detected.")
            log_message("❌ Internet: Not connected")

    elif "joke" in command:
        joke = pyjokes.get_joke()
        speak(joke)
        log_message(f"😂 {joke}")

    elif 'exit' in command or 'quit' in command or 'goodbye' in command:
        speak("Goodbye!")
        root.quit()

    else:
        # Use Gemini AI for unknown commands
        if GEMINI_ENABLED:
            log_message(f"🤖 Asking Gemini AI...")
            update_status("🤖 Thinking...", "#9c27b0")
            
            ai_response = ask_gemini(command)
            
            # Limit response length for speech
            if len(ai_response) > 300:
                speech_response = ai_response[:300] + "... Check the log for full response."
            else:
                speech_response = ai_response
            
            speak(speech_response)
            log_message(f"🤖 Gemini: {ai_response}")
            update_status("🟢 Listening...", "#00ff88")
        else:
            speak("I don't understand that command. Please try again.")
            log_message("❓ Unknown command")


def start_assistant():
    threading.Thread(target=listen, daemon=True).start()
    speak("Voice assistant is active.")
    update_status("🟢 Active", "#00ff88")
    log_message("🎤 Voice Assistant Started")


def update_status(text, color):
    """Update status label"""
    status_label.config(text=text, foreground=color)


def log_message(message):
    """Add message to output box"""
    output_text.insert(tk.END, f"{message}\n")
    output_text.see(tk.END)


def show_help():
    help_text = """
🎤 VOICE COMMANDS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📁 FILE OPERATIONS
• create file [name]
• create folder [name]
• rename file [old] to [new]
• rename folder [old] to [new]

🖥️ SYSTEM CONTROLS
• open [app name]
• shutdown
• lock computer
• take screenshot

🔊 VOLUME & BRIGHTNESS
• set volume to [0-100]
• increase volume / decrease volume
• set brightness to [0-100]

🌐 INTERNET & SEARCH
• search for [query]
• start youtube
• wikipedia [topic]

🖱️ WINDOW CONTROLS
• minimize window
• maximize window
• scroll up / scroll down

⌨️ TYPING
• type [your text]

📊 SYSTEM INFO
• battery status
• check internet
• what's the date
• what's the time

🤖 AI ASSISTANT
• Ask any question and Gemini AI will answer!
  Examples:
  - "explain quantum physics"
  - "write a poem about technology"
  - "what's the capital of France"

😄 FUN
• tell me a joke
    """
    
    help_window = tk.Toplevel(root)
    help_window.title("📚 Help - Voice Commands")
    help_window.geometry("600x700")
    help_window.configure(bg="#1a1a2e")
    
    help_frame = tk.Frame(help_window, bg="#1a1a2e")
    help_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    help_text_widget = scrolledtext.ScrolledText(
        help_frame,
        wrap=tk.WORD,
        font=("Consolas", 10),
        bg="#16213e",
        fg="#00ff88",
        insertbackground="#00ff88",
        relief=tk.FLAT,
        padx=15,
        pady=15
    )
    help_text_widget.pack(fill=tk.BOTH, expand=True)
    help_text_widget.insert(tk.END, help_text)
    help_text_widget.config(state=tk.DISABLED)


def configure_gemini():
    """Open dialog to configure Gemini API key"""
    config_window = tk.Toplevel(root)
    config_window.title("⚙️ Configure Gemini API")
    config_window.geometry("500x200")
    config_window.configure(bg="#1a1a2e")

    tk.Label(
        config_window,
        text="🔑 Enter Your Gemini API Key",
        font=("Segoe UI", 14, "bold"),
        fg="#00ff88",
        bg="#1a1a2e"
    ).pack(pady=20)

    api_entry = tk.Entry(
        config_window,
        font=("Consolas", 11),
        bg="#16213e",
        fg="#ffffff",
        insertbackground="#00ff88",
        relief=tk.FLAT,
        width=50
    )
    api_entry.pack(pady=10, padx=20)
    api_entry.insert(0, GEMINI_API_KEY if GEMINI_API_KEY != "YOUR_GEMINI_API_KEY_HERE" else "")

    def save_api_key():
        global GEMINI_API_KEY, GEMINI_ENABLED, client

        api_key = api_entry.get().strip()
        if api_key:
            try:
                # New Gemini client (correct)
                client = genai.Client(api_key=api_key)

                GEMINI_API_KEY = api_key
                GEMINI_ENABLED = True

                messagebox.showinfo("Success", "✅ Gemini API configured successfully!")
                log_message("✅ Gemini AI is now active")
                config_window.destroy()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to configure API: {str(e)}")
        else:
            messagebox.showwarning("Warning", "Please enter an API key")

    tk.Button(
        config_window,
        text="💾 Save API Key",
        command=save_api_key,
        font=("Segoe UI", 11, "bold"),
        bg="#00ff88",
        fg="#1a1a2e",
        activebackground="#00cc70",
        relief=tk.FLAT,
        cursor="hand2",
        padx=20,
        pady=10
    ).pack(pady=20)

    tk.Label(
        config_window,
        text="Get your free API key at: https://aistudio.google.com",
        font=("Segoe UI", 9),
        fg="#888888",
        bg="#1a1a2e"
    ).pack()

    
    tk.Label(
        config_window,
        text="Get your free API key at: https://makersuite.google.com/app/apikey",
        font=("Segoe UI", 9),
        fg="#888888",
        bg="#1a1a2e"
    ).pack()


# Create main UI window
root = tk.Tk()
root.title("🎤 AI Voice Assistant")
root.geometry("800x600")
root.configure(bg="#0f0f23")
root.resizable(False, False)

# Header Frame
header_frame = tk.Frame(root, bg="#1a1a2e", height=80)
header_frame.pack(fill=tk.X)
header_frame.pack_propagate(False)

title_label = tk.Label(
    header_frame,
    text="🤖 AI VOICE ASSISTANT",
    font=("Segoe UI", 24, "bold"),
    fg="#00ff88",
    bg="#1a1a2e"
)
title_label.pack(pady=20)

# Status Frame
status_frame = tk.Frame(root, bg="#16213e", height=50)
status_frame.pack(fill=tk.X, padx=20, pady=(10, 0))
status_frame.pack_propagate(False)

status_label = tk.Label(
    status_frame,
    text="⚪ Ready",
    font=("Segoe UI", 12, "bold"),
    fg="#888888",
    bg="#16213e"
)
status_label.pack(pady=12)

# Control Panel Frame
control_frame = tk.Frame(root, bg="#0f0f23")
control_frame.pack(pady=20)

button_style = {
    "font": ("Segoe UI", 11, "bold"),
    "relief": tk.FLAT,
    "cursor": "hand2",
    "width": 18,
    "height": 2
}

# Start Button
start_button = tk.Button(
    control_frame,
    text="▶️ START ASSISTANT",
    command=start_assistant,
    bg="#00ff88",
    fg="#0f0f23",
    activebackground="#00cc70",
    **button_style
)
start_button.grid(row=0, column=0, padx=10, pady=5)

# Stop Button
stop_button = tk.Button(
    control_frame,
    text="⏹️ STOP LISTENING",
    command=stop_listening,
    bg="#ff9800",
    fg="#0f0f23",
    activebackground="#e68900",
    **button_style
)
stop_button.grid(row=0, column=1, padx=10, pady=5)

# Help Button
help_button = tk.Button(
    control_frame,
    text="❓ HELP",
    command=show_help,
    bg="#2196f3",
    fg="#ffffff",
    activebackground="#1976d2",
    **button_style
)
help_button.grid(row=1, column=0, padx=10, pady=5)

# Configure Gemini Button
gemini_button = tk.Button(
    control_frame,
    text="🤖 CONFIGURE AI",
    command=configure_gemini,
    bg="#9c27b0",
    fg="#ffffff",
    activebackground="#7b1fa2",
    **button_style
)
gemini_button.grid(row=1, column=1, padx=10, pady=5)

# Exit Button
exit_button = tk.Button(
    control_frame,
    text="❌ EXIT",
    command=root.quit,
    bg="#f44336",
    fg="#ffffff",
    activebackground="#d32f2f",
    **button_style
)
exit_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

# Output Frame
output_frame = tk.Frame(root, bg="#0f0f23")
output_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))

tk.Label(
    output_frame,
    text="📋 Activity Log",
    font=("Segoe UI", 12, "bold"),
    fg="#00ff88",
    bg="#0f0f23"
).pack(anchor=tk.W, pady=(0, 5))

output_text = scrolledtext.ScrolledText(
    output_frame,
    wrap=tk.WORD,
    font=("Consolas", 10),
    bg="#16213e",
    fg="#e0e0e0",
    insertbackground="#00ff88",
    relief=tk.FLAT,
    padx=15,
    pady=15
)
output_text.pack(fill=tk.BOTH, expand=True)

# Initial log message
log_message("🚀 AI Voice Assistant Ready!")
log_message("💡 Click 'START ASSISTANT' to begin")
if not GEMINI_ENABLED:
    log_message("⚠️ Configure Gemini API for AI features")

# Start Tkinter loop

root.mainloop()

