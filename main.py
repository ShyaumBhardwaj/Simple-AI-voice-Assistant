import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import datetime
import tkinter as tk
from tkinter import scrolledtext

# Initialize API key and speaker
apikey = 'your_openai_api_key'
openai.api_key = apikey
speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""

def chat(query):
    global chatStr
    chatStr += f"Harry: {query}\nJarvis: "
    try:
        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=chatStr,
            max_tokens=256,
            temperature=0.7,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        response_text = response.choices[0].text.strip()
        say(response_text)
        chatStr += response_text + "\n"
    except Exception as e:
        say("I encountered an error while processing your request.")
        print(f"Error with OpenAI service: {e}")

def say(text):
    speaker.Speak(text)
    display_conversation(text, "Jarvis")

def display_conversation(text, speaker):
    chat_window.configure(state='normal')
    chat_window.insert(tk.END, f"{speaker}: {text}\n")
    chat_window.configure(state='disabled')
    chat_window.yview(tk.END)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        display_conversation("Listening...", "System")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            display_conversation(query, "Harry")
            return query.lower()
        except Exception as e:
            display_conversation("I didn't catch that. Please repeat.", "System")
            print("Recognition error:", e)
            return None

def handle_command():
    query = takeCommand()
    if query:
        process_query(query)

def process_query(query):
    if "play" in query and "music" in query:
        song_name = query.replace("play", "").replace("music", "").strip()
        play_music_on_youtube(song_name)
    elif "the time" in query:
        now = datetime.datetime.now()
        say(f"Sir, the time is {now.strftime('%H:%M')}")
    elif "open" in query:
        for site in [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"]]:
            if site[0] in query:
                webbrowser.open(site[1])
                say(f"Opening {site[0]} sir...")
    elif "jarvis quit" in query:
        say("Goodbye.")
        root.quit()
    elif "reset chat" in query:
        global chatStr
        chatStr = ""
        chat_window.configure(state='normal')
        chat_window.delete('1.0', tk.END)
        chat_window.configure(state='disabled')
    else:
        chat(query)

def play_music_on_youtube(song_name):
    query_string = "+".join(song_name.split())
    url = f"https://www.youtube.com/results?search_query={query_string}"
    webbrowser.open(url)
    say(f"Playing {song_name} on YouTube")

# Create main window
root = tk.Tk()
root.title("Jarvis AI")
root.geometry("600x400")

# Create widgets
chat_window = scrolledtext.ScrolledText(root, state='disabled', height=20, width=70)
chat_window.pack(pady=10)

listen_button = tk.Button(root, text="Listen", command=handle_command)
listen_button.pack(pady=5)

# Start the GUI
root.mainloop()
say("Hello, Welcome to JarvisAI, How may I help")
