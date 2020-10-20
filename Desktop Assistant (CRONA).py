import os  
import random
import webbrowser
from builtins import int, str
from datetime import datetime
import playsound  
import speech_recognition as sr
import wolframalpha
import smtplib
from gtts import gTTS  
import cv2
import pyttsx3
from tkinter import *
import PIL.Image, PIL.ImageTk
from SongModuleMain import songModuleMain
from FileSearch import Search_File

global var,var1
var = StringVar()
var1 = StringVar()

window = Tk()


def wishMe():
    time = int(datetime.now().hour)
    if 0 <= time < 12:
        var.set("Good Morning")
        window.update()
        assistant_speaks("Good Morning")

    if 12 <= time < 16:
        var.set("Good Afternoon")
        window.update()
        assistant_speaks("Good Afternoon")

    if 16 <= time < 20:
        var.set("Good Evening")
        window.update()
        assistant_speaks("Good Evening")

    if 20 <= time < 0:
        var.set("Good Night")
        window.update()
        assistant_speaks("Good Night")



num = 1


def assistant_speaks(output):
    global num

    # num to rename every audio file
    # with different name to remove ambiguity
    num += 1
    print("Crona : ", output)

    toSpeak = gTTS(text=output, lang='en', slow=False)
    # saving the audio file given by google text to speech
    file = str(num) + ".mp3"
    toSpeak.save(file)

    # playsound package is used to play the same file.
    playsound.playsound(file, True)
    os.remove(file)


def get_audio():
    rObject = sr.Recognizer()
    audio = ' '

    with sr.Microphone() as source:
        var.set("speak...")
        window.update()
        print("Speak...")

        # recording the audio using speech recognition
        audio = rObject.listen(source, phrase_time_limit=4)
        var.set("stop...")
        window.update()
    print("Stop...")  # limit 2 secs

    try:

        text = rObject.recognize_google(audio, language='en-US')
        print("You : ", text)
        return text

    except:

        assistant_speaks("Could not understand your audio, PLease try again !")
        return 0


def play_music():
    var.set("Okay, here is your music! Enjoy!")
    window.update()
    assistant_speaks('Okay, here is your music! Enjoy!')
    music_folder = "E:\music\Music\\"
    songs = os.listdir(music_folder)
    music = random.randint(0, 27)
    os.startfile(os.path.join(music_folder, songs[music]))
    return 0


def word_open():
    var.set("Opening Microsoft Word")
    window.update()
    assistant_speaks("Opening Microsoft Word")
    os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\\Microsoft Office Word 2007.lnk')
    return 0


def paint_open():
    var.set("Opening paint")
    window.update()
    assistant_speaks("Opening paint")
    os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\\Paint.lnk')
    return 0


def excel_open():
    var.set("Opening excel")
    window.update()
    assistant_speaks("Opening excel")
    os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\\Microsoft Office Excel 2007.lnk')
    return 0


def powerpoint_open():
    var.set("Opening powerpoint")
    window.update()
    assistant_speaks("Opening powerpoint")
    os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\\Microsoft Office PowerPoint 2007.lnk')
    return 0


def youtube_open():
    var.set("okay! opening youtube...")
    window.update()
    assistant_speaks("okay! opening youtube")
    webbrowser.open('https://www.youtube.com/')
    return 0


def google_open():
    var.set("okay! opening google...")
    window.update()
    assistant_speaks("okay! opening google")
    webbrowser.open('https://www.google.com/')
    return 0


def define():
    var.set('Hello My name is Crona.I am desktop assistant.')
    window.update()
    assistant_speaks('Hello My name is Crona.I am desktop assistant.')
    var.set('I am design to perform some predefined tasks.')
    window.update()
    assistant_speaks('I am design to perform some predefined tasks.')
    var.set('My current tasks are like...')
    window.update()
    assistant_speaks('My current tasks are like...')
    var.set('application launch')
    window.update()
    assistant_speaks('application launch')
    var.set('web search')
    window.update()
    assistant_speaks('web search')
    var.set('play music')
    window.update()
    assistant_speaks('play music')
    var.set('calculation')
    window.update()
    assistant_speaks('calculation')
    var.set('mail send')
    window.update()
    assistant_speaks('mail send')
    var.set('location search')
    window.update()
    assistant_speaks('location search')
    var.set('file search')
    window.update()
    assistant_speaks('file search')
    return 0


def send_mail(msg):
    
    assistant_speaks('establishing secure connection...')
    connection = smtplib.SMTP('smtp.gmail.com', 587)   
    connection.ehlo()
    connection.starttls()
    assistant_speaks('connection established')
    assistant_speaks('logging into your gmail account')
    connection.login('cronaproject.12@gmail.com', '987crona123')
    assistant_speaks('logged in successfully')
    assistant_speaks('sending your mail...')
    connection.sendmail('cronaproject.12@gmail.com', 'gs972745@gmail.com',msg)
    var.set('mail sent successfully')
    window.update()
    assistant_speaks('mail sent successfully')
    connection.quit()
    return 0


def play():
    wishMe()
    btn1.configure(bg='orange')
    while True:
        var.set("What can i do for you?")
        window.update()
        assistant_speaks("What can i do for you?")
        input = get_audio().lower()
        var1.set(input)
        window.update()
        if 'open youtube' in input:
            youtube_open()
            break

        elif 'search' in input:
            var.set('okay! searching...')
            window.update()
            assistant_speaks('okay! searching...')
            ur = "https://www.google.co.in/search?q="
            webbrowser.open(ur + input)
            break

        elif "open google" in input or 'google' in input:
            google_open()
            break

        elif 'play music' in input or 'next track' in input:
            play_music()
            break

        elif 'define' in input:
            define()
            break

        elif 'locate' in input:
            ur = "https://www.google.com/maps/search/?api=1&query="
            var.set('okay!locating...')
            window.update()
            assistant_speaks('okay!locating...')
            webbrowser.open(ur + input)
            break

        elif "calculate" in input.lower():
            app_id = "G62579-HYUE6V3ETY"
            client = wolframalpha.Client(app_id)
            indx = input.lower().split().index('calculate')
            query = input.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            var.set("The answer is " + answer)
            window.update()
            assistant_speaks("The answer is " + answer)
            break

        elif "open word" in input:
            word_open()
            break

        elif "open paint" in input:
            paint_open()
            break

        elif "open excel" in input:
            excel_open()
            break

        elif "open powerpoint" in input:
            powerpoint_open()
            break

        elif "send mail" in input or "mail" in input:
            var.set('record your message...')
            window.update()
            assistant_speaks('record your message')
            msg = get_audio().lower()
            send_mail(msg)
            break

        elif "exit" in str(input) or "bye" in str(input):
            var.set("Ok bye,Have a good day!")
            window.update()
            assistant_speaks("Ok bye,Have a good day!")
            break

        elif "find" or 'file' in input:
            Search_File()
            break

        else:
            assistant_speaks('command not found,Next Command please...')






def update(ind):
    frame = frames[(ind)%100]
    ind += 1
    label.configure(image=frame)
    window.after(100, update, ind)
    window.configure(background='black')

label2 = Label(window, textvariable = var1, bg = '#ADD8E6')
label2.config(font=("Courier", 20))
var1.set('User command')
label2.pack()

label1 = Label(window, textvariable = var, bg = '#ADD8E6')
label1.config(font=("Courier", 20))
var.set('Welcome')
label1.pack()

frames = [PhotoImage(file='Assistant.gif',format = 'gif -index %i' %(i)) for i in range(100)]
window.title('Crona (Desktop Assistant)')

label = Label(window, width = 500, height = 500)
label.pack()
window.after(0, update, 0)

btn0 = Button(text = 'ABOUT',width = 20, command = define, bg = '#5C85FB')
btn0.config(font=("Courier", 12))
btn0.pack()
btn1 = Button(text = 'PLAY',width = 20,command = play, bg = '#5C85FB')
btn1.config(font=("Courier", 12))
btn1.pack()
btn2 = Button(text = 'EXIT',width = 20, command = window.destroy, bg = '#5C85FB')
btn2.config(font=("Courier", 12))
btn2.pack()

window.mainloop()

