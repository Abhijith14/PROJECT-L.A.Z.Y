import os
import pandas
import datetime
from datetime import date, datetime
from time import sleep
import subprocess
import pyautogui
import pyttsx3

period_check = 0

week_list = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']
time_table = []
nextsub_time = -1
nextsub_dayorder = -1
day_now = ''
time_now = ''
time_main_list = ''
time_main_duration = ''
no_of_periods = ''
dayorder_no = ''
day_list = []
subject_list = ''
app_list = ''
link_list = ''
meetid_list = ''
pass_list = ''
app_dest_list = ''
name_list = ''

curr_dayorder = ''

def Talk(speech):
    engine = pyttsx3.init()
    voice = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0"
    engine.setProperty('voice', voice)
    engine.say(speech)
    engine.runAndWait()



def get_data():
    global name_list, app_dest_list, time_table, time_main_list, time_main_duration, no_of_periods, dayorder_no, day_list, subject_list, app_list, link_list, meetid_list, pass_list, app_dest_list
    df1 = pandas.read_excel('TimeTable/TT.xlsx', sheet_name="TIME")
    df2 = pandas.read_excel('TimeTable/TT.xlsx', sheet_name="DATA")
    df3 = pandas.read_excel('TimeTable/TT.xlsx', sheet_name="DAYS")
    df4 = pandas.read_excel('TimeTable/TT.xlsx', sheet_name="LINK")

    time_main_list = df1['TIME'].tolist()
    time_main_duration = df1['DURATION'].tolist()


    no_of_periods = len(time_main_list)

    for i in time_main_list:
        time_table.append(df2[i].tolist())

    dayorder_no = df2['DAYORDER'].tolist()

    dayorder_temp = df3['DAYORDER'].tolist()
    dayorder_day = df3['DAY'].tolist()

    for i in dayorder_no:
        day_list.append(dayorder_day[dayorder_temp.index(i)])


    subject_list = df4['SUBJECT'].tolist()
    app_list = df4['APP'].tolist()
    link_list = df4['LINK'].tolist()
    meetid_list = df4['MEET_ID'].tolist()
    pass_list = df4['PASSWORD'].tolist()
    app_dest_list = df4['APP_DEST'].tolist()
    name_list = df4['NAME'].tolist()


def curr_status():
    global day_now, time_now, week_list
    today = date.today()
    now = datetime.now()
    # today = datetime.datetime(2020, 11, 9)
    week_no = today.weekday()
    day_now = week_list[week_no]
    time_now = now#.strftime("%H.%M")


def check_data():
    global curr_dayorder, no_of_periods, dayorder_no, day_now, time_now, no_of_periods, period_check
    get_data()
    curr_status()
    time_list = []
    time_sub_list_start = []
    time_sub_list_end = []
    S_hrs = []
    S_mins = []
    E_hrs = []
    E_mins = []
    curr_dayorder = dayorder_no[day_list.index(day_now)]

    for i in time_main_duration:
        time_list.append(i.split(' to '))

    for i in range(int(no_of_periods)):
        start = time_list[i][0]
        end = time_list[i][1]
        time_sub_list_start.append(start.split('.'))
        time_sub_list_end.append(end.split('.'))

    for i in time_sub_list_start:
        S_hrs.append(i[0])
        S_mins.append(i[1])
    for i in time_sub_list_end:
        E_hrs.append(i[0])

        E_mins.append(i[1])

    # print(time_now)

    for i in range(int(no_of_periods)):
        Sh = time_sub_list_start[i][0]
        Sm = time_sub_list_start[i][1]
        Eh = time_sub_list_end[i][0]
        Em = time_sub_list_end[i][1]
        StartTime = time_now.replace(hour=int(Sh), minute=int(Sm), second=0, microsecond=0)
        EndTime = time_now.replace(hour=int(Eh), minute=int(Em), second=0, microsecond=0)
        # StartTime = time_now.replace(hour=int(0), minute=int(0), second=0, microsecond=0)
        # EndTime = time_now.replace(hour=int(1), minute=int(47), second=0, microsecond=0)
        if StartTime <= time_now <= EndTime:
            print("Detected period " + str(i+1))
            period_check = i
            join_meeting(day_list.index(day_now), i, EndTime)
            break
        # else:
        #     print("Not " + str(i+1) + " period")


def closegc():
    sleep(1)
    print("Closing GC")
    subprocess.call(["taskkill", "/F", "/IM"    , "chrome.exe"])
    print("Searching for Next Period ...")
    sleep(2)

def closezoom():
    sleep(1)
    print("Closing ZOOM")
    end = pyautogui.locateCenterOnScreen('assets/Zend.png')
    sleep(0.5)
    pyautogui.moveTo(end)
    sleep(0.5)
    pyautogui.click(end)
    sleep(0.5)

    subprocess.call(["taskkill", "/F", "/IM", "Zoom.exe"])
    print("Searching for Next Period ...")
    sleep(2)

def opengc(link, app_dest_main):
    global period_check
    found = 0
    print("Opening GC")
    subprocess.Popen([app_dest_main, link])
    sleep(5)
    start = pyautogui.locateCenterOnScreen('assets/join.png')
    if start is None:
        numb = 0
        while start is None:
            numb = numb + 1
            start = pyautogui.locateCenterOnScreen('assets/join.png')
            print("rotation " + str(numb))
            print(start)
            if numb == 10:
                opengc(link, app_dest_main)
                found = 1
                break
    else:
        print("Found")
    if found == 0:
        sleep(0.5)
        pyautogui.hotkey("ctrl", "d")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "e")
        sleep(1)
        pyautogui.moveTo(start)
        sleep(1)
        print(start)
        pyautogui.click(start)
        period_check = period_check + 1


def openzoom(user, passw, dest, stud_name):
    global period_check
    print("Opening ZOOM")
    subprocess.Popen(dest)
    sleep(10)
    start = pyautogui.locateCenterOnScreen('assets/Zjoin.png')
    if start is None:
        print("Coudnt Start the Class")
        # openzoom()
    else:
        print(start)
        pyautogui.moveTo(start)
        sleep(1)
        pyautogui.click(start)
        sleep(5)
        pyautogui.typewrite(str(user))
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "a")
        sleep(0.5)
        pyautogui.keyDown("delete")
        sleep(0.5)
        pyautogui.keyUp("delete")
        sleep(0.5)
        pyautogui.typewrite(stud_name)
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.keyDown("enter")
        sleep(0.5)
        pyautogui.keyUp("enter")
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.keyDown("enter")
        sleep(0.5)
        pyautogui.keyUp("enter")
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.keyDown("enter")
        sleep(0.5)
        pyautogui.keyUp("enter")
        sleep(5)
        pyautogui.typewrite(str(passw))
        sleep(0.5)
        pyautogui.keyDown("tab")
        sleep(0.5)
        pyautogui.keyUp("tab")
        sleep(0.5)
        pyautogui.keyDown("enter")
        sleep(0.5)
        pyautogui.keyUp("enter")
        sleep(0.5)
        audio = pyautogui.locateCenterOnScreen('assets/audio.png')
        while audio is None:
            print("Checking if admitted...")
            audio = pyautogui.locateCenterOnScreen('assets/audio.png')
            sleep(0.5)
        pyautogui.moveTo(audio)
        sleep(1)
        pyautogui.click(audio)
        period_check = period_check + 1

def startmeeting(app, link, user, passw, app_dest_main, stud_name):
    global period_check
    if app == "GC":
        opengc(link, app_dest_main)
    elif app == "ZOOM":
        openzoom(int(user), passw, app_dest_main, stud_name)
    else:
        period_check = period_check + 1


def endmeeting(app, link, user, passw, EndTime):
    global time_now
    print(EndTime)
    print(time_now)

    while EndTime >= time_now:
        curr_status()
        sleep(1)
        print()
        print(EndTime)
        print(time_now)
        print()
    print("Finished!")
    if EndTime <= time_now:
        print("Time to leave !")
        print("Leaving in 10")
        Talk("Leaving in 10")
        for i in range(9):
            sleep(1)
            print(9-i)
            Talk(9-i)

        if app == "GC":
            pyautogui.hotkey("ctrl", "r")
            print("Closing in 5")
            for i in range(4):
                sleep(1)
                print(4 - i)
            closegc()
        elif app == "ZOOM":
            pyautogui.hotkey("alt", "f4")
            print("Closing in 5")
            for i in range(4):
                sleep(1)
                print(4 - i)
            closezoom()
        else:
            print("Ended")


def join_meeting(dayorder, timeperiod, EndTime):
    global time_table, app_list, subject_list, meetid_list, pass_list, nextsub_time, nextsub_dayorder, name_list
    nextsub_time = timeperiod
    nextsub_dayorder  = dayorder
    sub = time_table[timeperiod][dayorder]
    app = app_list[subject_list.index(time_table[timeperiod][dayorder])]
    app_dest_main = app_dest_list[subject_list.index(time_table[timeperiod][dayorder])]
    link = link_list[subject_list.index(time_table[timeperiod][dayorder])]
    user = meetid_list[subject_list.index(time_table[timeperiod][dayorder])]
    passw = pass_list[subject_list.index(time_table[timeperiod][dayorder])]
    stud_name = name_list[subject_list.index(time_table[timeperiod][dayorder])]
    print(sub)
    # print(time_table[timeperiod+5][dayorder])
    startmeeting(app, link, user, passw, app_dest_main, stud_name)
    print("Joined " + sub)
    Talk("Joined " + sub)
    endmeeting(app, link, user, passw, EndTime)


while True:
    check_data()
    if period_check == no_of_periods:
        break
    else:
        print("Searching for period "+str(period_check+1))
        if nextsub_dayorder == -1:
            print("Starting Class. Pls wait...")
        else:
            print("Subject : " + time_table[nextsub_time+1][nextsub_dayorder])
        sleep(1)
    print()