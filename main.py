import win32com.client as win32
from tkinter.filedialog import askopenfile
import tkinter as tk
import pandas as pd

emails = []
names = []
global body_template 
body_template = ["Hello, ", "$name", "\n", "We want to inform you that you have an unpaid invoice with our company. You can find the details of outstanding payment bellow:"]
# The user interface
class MyGui:
    def __init__(self, master):
        
        self.root = master
        self.root.geometry("600x600")

        # Header
        header = tk.Label(self.root, text="Bulk email sender", font=("Arial", 18))
        header.pack(padx=20, pady=20)

        templateLabel = tk.Label(self.root, text="Email Template", font=("Arial", 12))
        templateLabel.pack()

        editTemplateLabel = tk.Label(self.root, text="You can edit the template bellow. you need to write $name instead of the client name.", font=("Arial", 8))
        editTemplateLabel.pack()

        templateText = tk.Text(self.root, font=("Arial", 10), height=3)
        displayable_body_template = ""
        for sequence in body_template:
                if sequence == "$name":
                    displayable_body_template = displayable_body_template + "$name"
                else:
                    displayable_body_template = displayable_body_template+sequence

        # TODO, change the body template when the text in the templateText widget changes

        templateText.insert("end-1c", displayable_body_template)
        templateText.pack(padx=10, pady=10)
        templateText.bind('<KeyPress>', update_body_template(templateText))

        xFrame = tk.Frame(self.root)
        xFrame.columnconfigure(0, weight=1)
        xFrame.columnconfigure(1,weight=1)

        namesListBox = tk.Listbox(xFrame)
        emailsListBox = tk.Listbox(xFrame)
        openFileButton = tk.Button(self.root, text="Open Excel file", command=lambda:self.chooseFile(namesListBox, emailsListBox))
        openFileButton.pack()

        namesListBox.grid(row=0, column =0, sticky= tk.W+tk.E)
        emailsListBox.grid(row=0, column =1, sticky= tk.W+tk.E)
        xFrame.pack()

        sendEmailsButton = tk.Button(self.root, text="Send Emails", command=self.on_send_emails_button_click)
        sendEmailsButton.pack()
    


    # Let the user upload an Excel file with columns for name, email
    def chooseFile(self, names_listbox, emails_listbox): 
        filename = askopenfile()
        print("We are attempting to read the file: \n", filename)
        print()

        
        email_list = pd.read_excel(filename.name)
        print("Email list = ", email_list)

        global names 
        names = email_list['name'].to_list()
        print("The names we identified in the file are:\n", names)
        print()
        change_text(names_listbox, names)
        
        global emails
        emails = email_list['email'].to_list()
        print("The emails we identified are:\n", emails)
        print()
        change_text(emails_listbox, emails)

        

    def send_emails(self, subjects, bodies):
        
        #outlook = win32.Dispatch('Outlook.Application')
        for i in range(len(emails)):
            # for every record create an email
         #   mail = outlook.CreateItem(0)
          #  mail.To = emails[i]
           # mail.Subject = subjects[i]
            #mail.Body = bodies[i]

            print("For the item number", i, "we have the following email")
            print(emails[i])
            print()
            print(subjects[i])
            print()
            print(bodies[i])
 
    # sending the email
           # mail.Send()
    
    def on_send_emails_button_click(self):
        # Create lists for email subjects and bodies with the same number of elements as the emails list
        subjects = ["Unpaid Invoice Notification" for i in range(len(emails))]
        bodies = ["" for i in range(len(emails))]

        # Populate the lists with subjects and bodies as needed
        for i, name in enumerate(names):

            # Create an email body using the template and replacing the name where needed
            body = ""
            for sequence in body_template:
                if sequence == "$name ":
                    body = body + name
                else:
                    body = body+sequence
            
            bodies[i] = body
        self.send_emails(subjects, bodies)

def change_text(text_widget, new_text):
    # Delete the current content for the text widget
    text_widget.delete("0", "end")
    # Insert the new text for the text widget
    for word in new_text:
        text_widget.insert('0', word)
        
def update_body_template(text_widget):
    global body_template
    body_template = [""]
    print(text_widget.get("1.0",'end-1c'))
    for sequence in text_widget.get("1.0",'end-1c').split():
        body_template.append(sequence + " ")
    print("AICI", body_template)


def main():
    root = tk.Tk()
    app = MyGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()