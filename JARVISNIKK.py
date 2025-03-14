# ==== Importing all the necessary libraries
import speech_recognition as sr
import pyttsx3
import webbrowser
import datetime
from tkinter import *
from PIL import ImageTk
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import pygetwindow as gw

# ==== Class Assistant
class assistance_gui:
    def __init__(self, root):
        self.root = root
        self.root.title("Voice Assistant")
        self.root.geometry('600x600')

        self.bg = ImageTk.PhotoImage(file="images/image1.png")
        bg = Label(self.root, image=self.bg).place(x=0, y=0)

        self.centre = ImageTk.PhotoImage(file="images/frame.jpg")
        left = Label(self.root, image=self.centre).place(x=100, y=100, width=400, height=400)

        # ====start button
        start = Button(self.root, text='START', font=("times new roman", 14), command=self.start_option).place(x=150, y=520)

        # ====close button
        close = Button(self.root, text='CLOSE', font=("times new roman", 14), command=self.close_window).place(x=350, y=520)

    # ==== start assistant
    def start_option(self):
        listener = sr.Recognizer()
        engine = pyttsx3.init()

        # ==== Voice Control
        def speak(text):
            engine.say(text)
            engine.runAndWait()

        # ====Default Start
        def start():
            # ==== Wish Start
            hour = int(datetime.datetime.now().hour)
            if hour >= 0 and hour < 12:
                wish = "Good Morning!"
            elif hour >= 12 and hour < 18:
                wish = "Good Afternoon!"
            else:
                wish = "Good Evening!"
            speak('Hello Sir,' + wish + ' I am your voice assistant. Please tell me how may I help you')
            # ==== Wish End

        # ==== Take Command
        def take_command():
            try:
                with sr.Microphone() as data_taker:
                    print("Say Something")
                    voice = listener.listen(data_taker)
                    instruction = listener.recognize_google(voice)
                    instruction = instruction.lower()
                    return instruction
            except:
                pass

        # ==== Run command
        def run_command():
            instruction = take_command()
            print(instruction)
            try:
                if 'who are you' in instruction:
                    speak('I am your personal voice Assistant')

                elif 'what can you do for me' in instruction:
                    speak('I can play songs, tell time, and help you go with wikipedia')

                elif 'current time' in instruction:
                    time = datetime.datetime.now().strftime('%I:%M %p')
                    speak('Current time is ' + time)

                elif 'open google' in instruction:
                    speak('Opening Google')
                    webbrowser.open('https://www.google.com')

                elif 'search google for' in instruction:
                    topic = instruction.replace('search google for', '').strip()
                    speak(f'Searching Google for {topic}')
                    webbrowser.open(f'https://www.google.com/search?q={topic}')

                elif 'open youtube' in instruction:
                    speak('Opening YouTube')
                    webbrowser.open('https://www.youtube.com')

                elif 'search youtube for' in instruction:
                    topic = instruction.replace('search youtube for', '').strip()
                    speak(f'Searching YouTube for {topic}')
                    webbrowser.open(f'https://www.youtube.com/results?search_query={topic}')

                elif 'play song' in instruction:
                    song = instruction.replace('play song', '').strip()
                    speak(f'Playing {song} on YouTube')
                    webbrowser.open(f'https://www.youtube.com/results?search_query={song}')

                elif 'open facebook' in instruction:
                    speak('Opening Facebook')
                    webbrowser.open('https://www.facebook.com')

                elif 'open python geeks' in instruction:
                    speak('Opening PythonGeeks')
                    webbrowser.open('https://www.pythongeeks.org')

                elif 'open linkedin' in instruction:
                    speak('Opening LinkedIn')
                    webbrowser.open('https://www.linkedin.com')

                elif 'open gmail' in instruction:
                    speak('Opening Gmail')
                    webbrowser.open('https://www.gmail.com')

                elif 'open stack overflow' in instruction:
                    speak('Opening Stack Overflow')
                    webbrowser.open('https://www.stackoverflow.com')

                elif 'open instagram' in instruction:
                    speak('Opening Instagram')
                    webbrowser.open('https://www.instagram.com')

                elif 'search wikipedia for' in instruction:
                    topic = instruction.replace('search wikipedia for', '').strip()
                    speak(f'Searching Wikipedia for {topic}')
                    webbrowser.open(f'https://en.wikipedia.org/wiki/{topic}')

                elif 'volume up' in instruction:
                    devices = AudioUtilities.GetSpeakers()
                    interface = devices.Activate(
                        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    volume = cast(interface, POINTER(IAudioEndpointVolume))
                    current_volume = volume.GetMasterVolumeLevelScalar()
                    volume.SetMasterVolumeLevelScalar(min(current_volume + 0.1, 1.0), None)
                    speak('Volume increased')

                elif 'volume down' in instruction:
                    devices = AudioUtilities.GetSpeakers()
                    interface = devices.Activate(
                        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    volume = cast(interface, POINTER(IAudioEndpointVolume))
                    current_volume = volume.GetMasterVolumeLevelScalar()
                    volume.SetMasterVolumeLevelScalar(max(current_volume - 0.1, 0.0), None)
                    speak('Volume decreased')

                elif 'mute' in instruction:
                    devices = AudioUtilities.GetSpeakers()
                    interface = devices.Activate(
                        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    volume = cast(interface, POINTER(IAudioEndpointVolume))
                    volume.SetMute(1, None)
                    speak('Muted')

                elif 'switch tab' in instruction:
                    windows = gw.getWindowsWithTitle('')
                    if windows:
                        windows[0].activate()
                        speak('Switched tab')

                elif 'shutdown' in instruction:
                    speak('I am shutting down')
                    self.close_window()
                    return False
                else:
                    speak('I did not understand, can you repeat again')
            except:
                speak('Waiting for your response')
            return True

        # ====Default Start calling
        start()

        # ====To run assistance continuously
        while True:
            if run_command():
                run_command()
            else:
                break

    # ==== Close window
    def close_window(self):
        self.root.destroy()

# ==== create tkinter window
root = Tk()

# === creating object for class
obj = assistance_gui(root)

# ==== start the gui
root.mainloop()