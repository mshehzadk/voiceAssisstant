import speech_recognition as sr
import os
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def say(text):
    speaker.speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            print(e)
            return "Some Error Occurred. Sorry from Jarvis"


if __name__ == '__main__':
    text = takeCommand()
    say(text)
