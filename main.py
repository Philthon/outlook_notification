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
    location_label.configure(bg=color)
    if color == "#FFFFFF":
        button.configure(bg="#E21239")
    if color == "#E21239":
        button.configure(bg="#FFFFFF")


def play_sound(window, meeting):
    # credits: Sound Effect by https://pixabay.com/users/universfield-28281460/?utm_source=link-attribution&amp;\
    # utm_medium=referral&amp;utm_campaign=music&amp;utm_content=143029
    sound_path = os.getcwd() + "\\sound.mp3"
    now = datetime.datetime.now().replace(tzinfo=None)
    meeting_start = meeting.Start
    meeting_start = datetime.datetime.strptime(
        str(meeting_start), "%Y-%m-%d %H:%M:%S%z"
    )
    meeting_start = meeting_start.replace(tzinfo=None)
    playsound(sound_path)
    if meeting_start < now + datetime.timedelta(minutes=1):
        playsound(sound_path)
        window.after(500, lambda: play_sound(window, meeting))
        colors = ["#FFFFFF", "#E21239"]
        color = colors[0]
        if window.cget("bg") == colors[0]:
            color = colors[1]
        if window.cget("bg") == colors[1]:
            color = colors[0]
        window.after(100, lambda: change_bg(window, color))


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

    global location_label
    location_label = Label(
        window, text=meeting.Subject, font=("Roboto", 12), background="white"
    )
    location_label.pack()

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

    window.after(500, lambda: play_sound(window, meeting))

    window.mainloop()


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)

now = datetime.datetime.now()
start = now + datetime.timedelta(minutes=1)
end = now + datetime.timedelta(minutes=25)
items = calendar.Items
items.IncludeRecurrences = True
items.Sort("[Start]")
items = items.Restrict(
    "[Start] >= '{0}' AND [Start] <= '{1}'".format(
        start.strftime("%m/%d/%Y %H:%M"), end.strftime("%m/%d/%Y %H:%M")
    )
)
global meeting
meeting = None
for item in items:
    meeting = item
    break

if meeting:
    icon_path = os.getcwd() + "\\icon.ico"

    threading.Thread(target=show_notification(meeting)).start()
