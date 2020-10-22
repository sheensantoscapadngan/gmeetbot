from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from pynput.keyboard import Key, Controller
import speech_recognition as sr
from datetime import datetime
from datetime import date
import calendar
import pandas as pd
import random

keyboard = Controller()
schedule = pd.read_excel('schedule.xlsx')
chat_responses = ['halo','yes yes yes','present xd']


def listen_to_name_call():
    r = sr.Recognizer()
    #print(sr.Microphone.list_microphone_names())
    mic = sr.Microphone(device_index=3)
    start_ms = int(round(time.time()*1000))
    name_call = ['shin','seen','chin','scene','shane']

    while True:

        curr_ms = int(round(time.time()*1000))
        minute_diff = int((curr_ms - start_ms)/(1000*60))
        if minute_diff >= 15: break

        try:
            with mic as source: audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                print(text)
                for name in name_call:
                    if name in text:
                        response = chat_responses[random.randint(0,len(chat_responses)-1)]
                        keyboard.type(response)
                        press_and_release(Key.enter)
                        break
            except: print("CANNOT RECOGNIZE SPEECH!")
        except:
            print("AN ERROR OCCURED")

def findDate():
    born = date.today().weekday()
    return (calendar.day_name[born])


def findTime():
    now = datetime.now().time()
    return (findDate(),now.hour,now.minute)

def press_and_release(key):
    keyboard.press(key)
    keyboard.release(key)

def open_link(link):
    options = Options()
	#edit this to your corresponding user data location for google chrome
    options.add_argument("user-data-dir=C:\\Users\\Sheen Capadngan\\AppData\\Local\\Google\\Chrome\\User Data\\")
    driver = webdriver.Chrome(options = options)
    driver.get(link)
    time.sleep(4)

    #DISABLE CAMERA AND MIC AND ENTER ROOM
    for i in range(0,7):
        press_and_release(Key.tab)
        if i in [3,6]: press_and_release(Key.enter)
        time.sleep(1)

    passed_everyone = False
    time.sleep(3)
    #TYPE IN PRESENT
    for i in range(0,40):
        press_and_release(Key.tab)
        elem = driver.switch_to.active_element
        class_name = elem.get_attribute("class")
		#the following block of strings are the respective ids for the people and message icon within gmeet so this could vary
        if "uArJ5e UQuaGc kCyAyd" in class_name and not passed_everyone: passed_everyone = True
        elif "uArJ5e UQuaGc kCyAyd" in class_name:
            press_and_release(Key.enter)
            time.sleep(1)
            break
        time.sleep(0.4)

    time.sleep(2)
    keyboard.type("present ma'am")
    time.sleep(1)
    press_and_release(Key.enter)

    keyboard.press(Key.ctrl)
    keyboard.press('e')
    keyboard.release(Key.ctrl)
    keyboard.release('e')

    listen_to_name_call()
    driver.close()

def lookup_schedule():
    current_time = findTime()
    row_count = len(schedule.index)
    print("CURRENT TIME:",current_time)
    for i in range(row_count):
        current_row = schedule.iloc[i]
        link = current_row['LINK']
        subject_time = current_row['TIME']
        day = current_row['DAY OF WEEK']
        hour = subject_time.hour
        minute = subject_time.minute
        look_up = (day,hour,minute)
        print("LOOKUP RESULT:",look_up)
        if current_time == look_up:
            print("EXECUTING")
            open_link(link)
    print('\n')

def background_process():
    lookup_schedule()
    time.sleep(1)

while True: background_process()
