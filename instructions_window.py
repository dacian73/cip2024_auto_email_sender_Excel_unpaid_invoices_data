import tkinter as tk

from constants import *

def instructionsPage(self):
        # Create a tkinter window with instructions on using the app
        win=tk.Tk()
        win.geometry(WINDOW_SIZE)
        win.title("Instructions")
        win.iconbitmap(HELP_ICON_PATH)
        label1 = tk.Label(win, text= "How can you use the app? ",font=('Arial bold', 18)).pack(pady=20)
        label2 = tk.Label(win, text= "This app allows you to send emails to a large number of users...",font=('Arial', 12)).pack(pady=20)
        #Make the window jump above all
        win.attributes('-topmost',True)
        win.mainloop()