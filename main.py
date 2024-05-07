import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import datetime


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
            model="text-davinci-003",  # Ensure this is the right model
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


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except Exception as e:
            print("Recognition error:", e)
            return "I didn't catch that."


def play_music_on_youtube(song_name):
    query_string = "+".join(song_name.split())
    url = f"https://www.youtube.com/results?search_query={query_string}"
    webbrowser.open(url)
    say(f"Playing {song_name} on YouTube")


def main():
    print('Welcome to Jarvis A.I.')
    say("Welcome to Jarvis A.I.")
    while True:
        query = takeCommand()
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
            break
        elif "reset chat" in query:
            chatStr = ""
        else:
            chat(query)


if __name__ == '__main__':
    main()
