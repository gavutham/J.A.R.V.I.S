import pyautogui
import pyttsx3
import wikipedia
import speech_recognition as sr
from pynput.keyboard import Controller, Key
import time
import os
import tkinter
from tkinter import messagebox, simpledialog
import threading
import random
import subprocess
import requests
from collections import Counter
import pywhatkit
import passwordManager
import pyautogui
from speech_recognition import WaitTimeoutError, UnknownValueError, Recognizer, Microphone


class Jarvis:
    def __init__(self, rate=155, voice="female"):
        self.engine = pyttsx3.init()
        self.property_changing(rate=rate, voice=voice, volume=1)
        self.keyboard = Controller()
        self.directory = ""
        self.dir_available = []

    def say(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def property_changing(self, **kwargs):
        """Use rate to change rate,
         use volume to change volume value ranges from 0 to 1 and
         use voice to change voice value male or female"""

        if "rate" in kwargs:
            self.engine.setProperty('rate', kwargs['rate'])

        if "volume" in kwargs:
            self.engine.setProperty('volume', kwargs['volume'])

        if "voice" in kwargs:
            voice = self.engine.getProperty('voices')
            if kwargs["voice"].lower() == "female":
                self.engine.setProperty('voice', voice[1].id)
            if kwargs["voice"].lower() == "male":
                self.engine.setProperty('voice', voice[0].id)
        return None

    def speech_rec(self):
        # return input("command: ")
        r = sr.Recognizer()
        with sr.Microphone() as source:
            self.say("listening")
            r.adjust_for_ambient_noise(source, duration=0.5)
            audio = r.listen(source)
            try:
                said = r.recognize_google(audio)
                print(said)
                return said
            except Exception as exc:
                print(exc)

    def search_wiki(self, query: str):
        try:
            return str(wikipedia.summary(query)).split("\n")[0]

        except wikipedia.exceptions.PageError:
            avail_title = wikipedia.search(query)
            if len(avail_title) == 0:
                return "No search results found sir, please try again."
            else:
                print(avail_title)
                self.say("please choose what you want")
                title = self.speech_rec()
                return str(wikipedia.summary(title)).split("\n")[0]

    @staticmethod
    def get_time_now():
        t = time.strftime("%I:%M %p")
        return t

    @staticmethod
    def send_mail():
        return "Sorry sir, sending mail functionality is on testing, it will be available soon sir."

    def minimize(self):
        with self.keyboard.pressed(Key.cmd_l):
            self.keyboard.press(Key.down)
            self.keyboard.release(Key.down)
        time.sleep(0.01)
        with self.keyboard.pressed(Key.cmd_l):
            self.keyboard.press(Key.down)
            self.keyboard.release(Key.down)

    def maximize(self):
        with self.keyboard.pressed(Key.cmd_l):
            self.keyboard.press(Key.up)
            self.keyboard.release(Key.up)

    def go_back(self):
        self.keyboard.press(Key.backspace)
        self.keyboard.release(Key.backspace)
        lis = self.directory.split("\\")
        del lis[-1]
        self.directory = "\\".join(lis)
        print(self.directory)

    def close(self):
        with self.keyboard.pressed(Key.alt):
            self.keyboard.press(Key.f4)
            self.keyboard.release(Key.f4)

    @staticmethod
    def shutdown():
        subprocess.call("shutdown /p", shell=True)
        return "shutting down sir. bye, bye"

    def sleep(self):
        self.say("sleeping")
        subprocess.call("rundll32.exe powrprof.dll, SetSuspendState Sleep", shell=True)

    @staticmethod
    def restart():
        subprocess.call("shutdown /r", shell=True)
        return "restarting"

    def open(self, direc):
        def show_message():
            directory_ = ", ".join(self.dir_available)
            print(directory_)
            tk = tkinter.Tk()
            tk.withdraw()
            messagebox.showinfo("Choose Directory", directory_)

        if "drive" in direc:
            drive = direc[0].upper()
            self.directory = f"{drive}:\\"

        else:
            all_dir = os.listdir(self.directory)

            if direc in all_dir:
                self.close()
                self.directory = os.path.abspath(os.path.join(self.directory, direc))

            else:
                index = 1
                dirs = []

                for directory in all_dir:
                    dirs.append(f"{directory} - {index}\n")
                    index += 1

                self.dir_available = dirs

                self.say("choose the thing you want to open sir")

                th = threading.Thread(target=show_message)
                th.start()

                choice = str(self.speech_rec())

                self.keyboard.press(Key.enter)
                self.keyboard.release(Key.enter)
                time.sleep(0.1)


                try:
                    choice = int(choice) - 1

                except ValueError:
                    if choice.lower().strip() == "one":
                        choice = 0
                    elif choice.lower().strip() == "tu":
                        choice = 1
                    elif choice.lower().strip() == "tree":
                        choice = 2
                self.close()
                req = all_dir[choice]
                self.directory = os.path.abspath(os.path.join(self.directory, req))

        os.startfile(self.directory)
        print(self.directory)
        return "opening sir"

    def check_weather(self, lat=9.303290, lon=77.444183):

        api_key = "3d948560591cc99d0d93c9242fa1af88"  # from website
        url = "https://api.openweathermap.org/data/3.0/onecall"
        parameters = {
            "lat": lat,
            "lon": lon,
            "exclude": "current,minutely,daily,alerts",
            "appid": api_key
        }

        response = requests.get(url=url, params=parameters).json()
        print(response)
        lis = []
        i = 0

        while i < 18:
            final = response['hourly'][i]['weather'][0]['description']
            lis.append(final)
            i += 1

        most, next_ = self.most_frequent(lis)

        return f"Mostly today will be a {most[0]} day, there is some possibility for {next_[0]}"

    @staticmethod
    def most_frequent(liss):
        occurrence_count = Counter(liss)
        return occurrence_count.most_common(2)

    @staticmethod
    def adjust_bright(level):

        command = '''powershell -Command "Get-Ciminstance -Namespace root/WMI -ClassName WmiMonitorBrightness | Select -ExpandProperty "CurrentBrightness""'''
        levl = subprocess.check_output(command, shell=True).decode()
        lvl = int((int(levl) / 70) * 100)

        change_lvl = int((level/70) * 100)
        final_lvl = lvl + change_lvl

        subprocess.call(f'''powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{final_lvl})''')

    @staticmethod
    def playonyt(query):
        pywhatkit.playonyt(topic=query)

    @staticmethod
    def play(what):
        # command to run music files
        subprocess.call(f"powershell -c (New-Object Media.SoundPlayer 'G:\python projects\own\project_jarvis\\{what}.wav').PlaySync();")

    @staticmethod
    def say_hi(who):
        return f"Hi {who}. how are you doing today {who}?"

    def passwordManager(self):
        entry = simpledialog.askstring("Authorization", "Access Password: ", show="*")
        if entry == os.environ.get("jarvisaccesscode"):
            password = "nothing"
            pyautogui.typewrite(password)
            pyautogui.press("enter")
            self.say("Access Granted")
        else:
            self.say("Access Denied")


    def run_jarvis(self):

        self.say("On your command sir.")

        while True:

            word = str(self.speech_rec())
            said = word.lower().strip()

            if "jarvis" in said or "Jarvis" in said:
                said = said.replace("jarvis", '').strip()
                said = said.replace("Jarvis", '').strip()

            if "hey" == said or "hello" == said:
                lis = ["Yes Sir, what can i do for you now.", "Yeah sir, ready to serve for you"]
                word = random.choice(lis)

            elif "let's rock" == said:
                word = "playing"
                self.play("theme")

            elif "introduce yourself" == said or "introduce about yourself" == said:
                self.play("introduc")
                word = "sir."

            elif "how are you" == said or "what about you" == said:
                self.say("Fine sir. Thanks for asking, how about you")
                self.speech_rec()
                word = "Good to hear that, sir"

            elif "exit" == said or "go to sleep" == said or "ok bye" == said or "bye bye" == said:
                self.say("Exiting sir")
                break

            elif "what is the time now" == said or "time now" == said:
                now = self.get_time_now()
                word = now

            elif "minimise" == said or "minimise the current window" == said:
                self.minimize()
                word = "minimised sir"

            elif "maximize" == said or "maximize the current window" == said:
                self.maximize()
                word = "maximised sir"

            elif "go back" == said:
                self.go_back()
                word = "got back"

            elif "i want to send a mail" == said:
                response = self.send_mail()
                word = response

            elif ("open" in said) and (said.split(" ")[0] == "open"):

                direct = said.replace("open", "")
                word = self.open(direct.strip())

            elif "close window" == said or "close the current window" == said:
                self.close()
                word = "closed sir"

            elif "say hi to" in said or "say hello to" in said:

                person = said.replace("say", '').replace("hi", '').replace("to", '').replace("hello", '').strip()
                print(person)
                word = self.say_hi(person)

            elif "search for " in said:
                self.say("searching")
                query = said.replace('search', '').replace('for', '').strip()
                word = self.search_wiki(query)

            elif "shutdown" == said or "system shutdown" == said:
                word = self.shutdown()

            elif ("on youtube" in said or "in youtube" in said) and "play" in said:
                topic = said.replace("in youtube", '').replace("on youtube", '').replace("play", '').strip()
                self.playonyt(topic)
                word = "playing"

            elif "what about weather" ==said or "what about weather here" ==said or "what about weather today" == said:
                word = self.check_weather()

            elif "put the system on sleep" == said or "system sleep" == said:
                self.sleep()

            elif "restart the system" == said or "system restart" == said:
                word = self.restart()

            elif ("increase brightness by" in said) or ("decrease brightness by" in said):
                lvl = said.replace("increase brightness by", '').replace("decrease brightness by", '').replace("points", '').strip()
                if "decrease" in said:
                    lvl = -(int(lvl))
                self.adjust_bright(int(lvl))
                word = "done"

            elif ("password for" in said) or ("what is the password for" in said):
                said.replace("password for", "").strip()
                said.replace("what is the password for", "").strip()
                self.say("The functionality needs authorization")
                self.passwordManager()
                word = ""

            else:
                word = "Sorry sir. Cant get you."

            self.say(word)

def activate():
    while True:
        r = Recognizer()
        with Microphone() as source:
            try:
                r.adjust_for_ambient_noise(source, duration=0.5)
                print("listening...")
                audio = r.listen(source, timeout=2)
                got = r.recognize_google(audio)

            except (WaitTimeoutError, UnknownValueError):
                print("error")
                continue

            said = str(got).strip().lower()
            print(said)

        if said == "jarvis" or said == "javed" or said == "david" or said == "hey jarvis" or said == "are you there" or said == "are you there jarvis":
            jarvis = Jarvis()
            jarvis.run_jarvis()
            audio, said, got = ('', '', '')


def greet():
    machine = Jarvis()
    time_ = machine.get_time_now()
    if "AM" in time_:
        greeting = "Good morning."
    elif int(time_.split(':')[0]) > 3:
        greeting = "Good Evening"
    else:
        greeting = "Good afternoon"
    machine.say(f"welcome back sir, {greeting}")
    del machine
