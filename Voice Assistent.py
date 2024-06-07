import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import subprocess
import time

# Initialize the recognizer and the text-to-speech engine
recognizer = sr.Recognizer()
tts_engine = pyttsx3.init()

# Function to make the assistant speak
def speak(text):
    tts_engine.say(text)
    tts_engine.runAndWait()

# Function to take a command from the user
def take_command():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        try:
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            return query
        except sr.UnknownValueError:
            print("Sorry, I didn't get that. Please repeat.")
            speak("Sorry, I didn't get that. Please repeat.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            speak(f"Could not request results; {e}")
            return None

# Function to execute commands
def execute_command(command):
    if 'wikipedia' in command:
        speak("Searching Wikipedia...")
        command = command.replace("wikipedia", "").strip()
        if command:
            try:
                results = wikipedia.summary(command, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:
                speak("There are multiple results for your query. Please be more specific.")
            except wikipedia.exceptions.PageError as e:
                speak("I couldn't find any results for your query.")
            except Exception as e:
                speak("An error occurred while searching Wikipedia.")
    elif 'open youtube' in command:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")
    elif 'open google' in command:
        speak("Opening Google")
        webbrowser.open("https://www.google.com")
    elif 'time' in command:
        str_time = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"The time is {str_time}")
    elif 'date' in command:
        str_date = datetime.datetime.now().strftime("%Y-%m-%d")
        speak(f"Today's date is {str_date}")
    elif 'open word' in command:
        speak("Opening Microsoft Word")
        word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        os.startfile(word_path)
    elif 'open excel' in command:
        speak("Opening Microsoft Excel")
        excel_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        os.startfile(excel_path)
    elif 'open powerpoint' in command:
        speak("Opening Microsoft PowerPoint")
        ppt_path = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
        os.startfile(ppt_path)
    elif 'open code' in command:
        speak("Opening Visual Studio Code")
        code_path = r"C:\Users\Sai Krishna\Documents\Sai Krishna Programmes\Learn Fusion\Voice Assistent\Voice Assistent.py"
        os.startfile(code_path)
    elif 'goodbye' in command or 'exit' in command or 'stop' in command:
        speak("Goodbye! Have a great day!")
        exit()
    else:
        speak("I am sorry, I can't do that yet.")

def main():
    speak("Hello, I am your voice assistant. How can I help you today?")
    while True:
        command = take_command()
        if command:
            execute_command(command.lower())

if __name__ == "__main__":
    main()