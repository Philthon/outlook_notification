import datetime
import os
import threading
import webbrowser
from tkinter import Button, Label, Tk, ttk

import win32com.client
from playsound import playsound


def open_meeting_link():
    if meeting.Location:
        webbrowser.open(meeting.Location)


def change_bg(window, color):
    window.configure(bg=color)
    subject_label.configure(bg=color)
    meeting_label.configure(bg=color)
    if color == "#FFFFFF":
        button.configure(bg="#E21239", fg="#FFFFFF")
    if color == "#E21239":
        button.configure(bg="#FFFFFF")


def handle_pop_up(window, meeting, play_audio):
    sound_path = os.getcwd() + "\\sound.mp3"
    now = datetime.datetime.now().replace(tzinfo=None)
    meeting_start = datetime.datetime.strptime(
        str(meeting.Start), "%Y-%m-%d %H:%M:%S%z"
    )
    meeting_start = meeting_start.replace(tzinfo=None)
    if meeting_start <= now + datetime.timedelta(minutes=1):
        if play_audio is True:
            playsound(sound_path)
        colors = ["#FFFFFF", "#E21239"]
        color = colors[0]
        if window.cget("bg") == colors[0]:
            color = colors[1]
        if window.cget("bg") == colors[1]:
            color = colors[0]
    window.after(100, lambda: change_bg(window, color))
    window.after(300000, lambda: window.destroy())


def show_notification(meeting):
    window = Tk()
    window.iconbitmap(icon_path)
    ttk.Style().theme_use("xpnative")
    window.title("Upcoming Meeting")
    window.iconbitmap(icon_path)
    window.attributes("-topmost", True)
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    screen_height = screen_height - 260
    screen_width = screen_width - 320
    window.geometry("300x150+{}+{}".format(screen_width, screen_height))
    window.resizable(False, False)
    window.configure(background="white")

    start_time = str(meeting.Start)[11:16]

    global subject_label
    subject_label = Label(
        window, text=start_time, font=("Roboto", 14, "bold"), background="white"
    )
    subject_label.pack(pady=10)

    global meeting_label
    location_text = meeting.Subject
    if len(location_text) > 35:
        location_text = location_text[:35] + "..."
    meeting_label = Label(
        window, text=location_text, font=("Roboto", 12), background="white"
    )
    meeting_label.pack()

    if meeting.Location:
        global button
        button = Button(
            window,
            text="Open Meeting",
            command=open_meeting_link,
            font=("Roboto", 12, "underline"),
        )
        button.config(
            relief="flat",
            overrelief="flat",
            background="#DEE1E3",
            foreground="#174CA1",
            borderwidth=5,
        )
        button.pack(pady=10)

    if play_audio is True:
        playsound(os.getcwd() + "\\sound.mp3")

    window.after(60000, lambda: handle_pop_up(window, meeting, play_audio))

    window.mainloop()


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)

now = datetime.datetime.now()
start = now + datetime.timedelta(minutes=1)
end = now + datetime.timedelta(minutes=25)
items = calendar.Items
items.IncludeRecurrences = True
items.Sort("[Start]")
items_future = items.Restrict(
    "[Start] >= '{0}' AND [Start] <= '{1}'".format(
        start.strftime("%m/%d/%Y %H:%M"), end.strftime("%m/%d/%Y %H:%M")
    )
)

items_current = items.Restrict(
    "[Start] <= '{0}' AND [End] >= '{1}'".format(
        now.strftime("%m/%d/%Y %H:%M"), now.strftime("%m/%d/%Y %H:%M")
    )
)

global play_audio
play_audio = True
for item_current in items_current:
    play_audio = False

global meeting
meeting = None
for item in items_future:
    meeting = item
    break

if meeting:
    icon_path = os.getcwd() + "\\phone.ico"
    threading.Thread(target=show_notification(meeting)).start()
