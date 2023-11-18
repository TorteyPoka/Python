import tkinter
import customtkinter
from pytube import YouTube


def startdownload():
    try:
        ytlink = link.get()
        ytobj = YouTube(ytlink)
        audio = ytobj.streams.get_audio_only()
        audio.download()
    except:
        print("Error")
    print("Downloaded")


# system settings
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# app frame
app = customtkinter.CTk()
app.geometry("500x500")
app.title("YouTube Downloader")

title = customtkinter.CTkLabel(app, text="Insert a youtube link: ", font=("Arial", 20))
title.pack(padx=10, pady=10)

url = tkinter.StringVar()
link = customtkinter.CTkEntry(app, width=350, hight=50, textvariable=url)
link.pack()

download = customtkinter.CTkButton(app, text="Download", command=startdownload)
download.pack(padx=10, pady=10)

# run app
app.mainloop()
