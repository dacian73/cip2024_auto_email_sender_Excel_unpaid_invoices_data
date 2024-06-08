import win32com.client as win32
from tkinter.filedialog import askopenfile
import tkinter as tk
import pandas as pd

# The user interface
class MyGui:
    def __init__(self, master):
        emails = []

        def get_emails(event):
            emails = list(emailsText.get("1.0",'end-1c').split("\n"))

            print("TEEEEEEEEEEEEEEEEEEEEEST", emails, "SUPPPPPPPPPPPPPPPEEEEEEEEEEERRRRRRRR")
        
        self.root = master
        self.root.geometry("600x600")

        # Header
        header = tk.Label(self.root, text="Bulk email sender")
        header.pack(padx=20, pady=20)

        xFrame = tk.Frame(self.root)
        xFrame.columnconfigure(0, weight=1)
        xFrame.columnconfigure(1,weight=1)

        namesText = tk.Listbox(xFrame)
        emailsText = tk.Listbox(xFrame)
        openFileButton = tk.Button(self.root, text="Open Excel file", command=lambda:self.chooseFile(namesText, emailsText))
        openFileButton.pack()

     

        namesText.grid(row=0, column =0, sticky= tk.W+tk.E)
        emailsText.grid(row=0, column =1, sticky= tk.W+tk.E)
        xFrame.pack()

        emailsText.bind('<KeyPress>', get_emails)

        sendEmailsButton = tk.Button(self.root, text="Send Emails", command=lambda:self.send_emails(emails, subjects, bodies))
        sendEmailsButton.pack()

    # Let the user upload an Excel file with columns for name, email
    def chooseFile(self, names_text_widget, emails_text_widget): 
        filename = askopenfile()
        print("We are attempting to read the file: \n", filename)
        print()

        
        email_list = pd.read_excel(filename.name)
        print("Email list = ", email_list)

        names = email_list['name'].to_list()
        print("The names we identified in the file are:\n", names)
        print()
        change_text(names_text_widget, names)
        
        emails = email_list['email'].to_list()
        print("The emails we identified are:\n", emails)
        print()
        change_text(emails_text_widget, emails)

    def send_emails(emails, subjects, bodies):
        outlook = win32.Dispatch('Outlook.Application')
 

        for i in range(len(emails)):
 
    # for every record create an email
            mail = outlook.CreateItem(0)
            mail.To = emails[i]
            mail.Subject = subjects[i]
            mail.Body = bodies[i]
 
    # sending the email
            mail.Send()

def change_text(text_widget, new_text):
    # Delete the current content for the text widget
    text_widget.delete("0", "end")
    # Insert the new text for the text widget
    for word in new_text:
        text_widget.insert('0', word)



def main():
    root = tk.Tk()
    app = MyGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()