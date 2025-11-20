🎙️ AI Voice Assistant (Python + Gemini 2.0 Flash)

A powerful and intelligent desktop voice assistant built with **Python**, **Tkinter**, and **Google Gemini 2.0 Flash**.  
This assistant can listen, understand, speak, perform system operations, and provide AI-powered responses — all from your computer.

---

## 🚀 Features

### 🔊 Voice Interaction
- Wake-word detection: **“assistant”**
- Listens automatically using microphone
- Responds with natural text-to-speech (pyttsx3)
- 15-second active listening window after wake word

### 🤖 AI-Powered Responses
- Uses **Google Gemini 2.0 Flash**
- Answers any question
- Generates smart replies
- Chat history visible in UI

### 🖥 System Control Commands
- Open applications  
- Adjust volume  
- Control brightness  
- Take screenshots  
- Search Wikipedia  
- Open websites  
- Tell jokes  
- Manage windows  
- Copy/Paste text  
- System information  

### 🪟 Modern Tkinter GUI
- Start/Stop assistant button  
- Live activity log  
- Status indicator  
- API configuration window  
- Dark theme with neon accents  

### 🌤 Weather Support
- Fetches real-time weather (optional API)

---

## 🛠️ Tech Stack

| Component | Technology |
|----------|------------|
| UI | Tkinter |
| Speech-to-Text | SpeechRecognition |
| Text-to-Speech | pyttsx3 |
| AI Model | Gemini 2.0 Flash (google-genai) |
| System Control | PyAutoGUI, PyCAW, pygetwindow |
| Utilities | psutil, requests, wikipedia |

---

## 📦 Installation

### 1️⃣ Clone the repository
```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
