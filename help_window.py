import tkinter as tk

from constants import *

def helpPage(self):
        # Create a tkinter window for "About" info
        win=tk.Tk()
        win.geometry(WINDOW_SIZE)
        win.title(HELP_WINDOW_TITLE)
        win.iconbitmap(HELP_ICON_PATH)
        label = tk.Label(win, text= "About the app! ",font=('Arial bold', 18)).pack(pady=20)
        label = tk.Label(win, text= "This app was created by dacian73 for the Code in Place 2024 final project.\nThe sourcecode is available at https://github.com/dacian73 ",font=('Arial', 12)).pack(pady=20)
        #Make the window jump above all
        win.attributes('-topmost',True)
        win.mainloop()