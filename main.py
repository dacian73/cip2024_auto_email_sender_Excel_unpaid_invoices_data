import win32com.client as win32
from tkinter.filedialog import askopenfile
import tkinter as tk
import pandas as pd

class MyGui:
    def __init__(self, master):
        self.root = master
        self.root.geometry("600x600")
        header = tk.Label(self.root, text="Bulk email sender")
        header.pack(padx=20, pady=20)

        xFrame = tk.Frame(self.root)
        xFrame.columnconfigure(0, weight=1)
        xFrame.columnconfigure(1,weight=1)

        namesText = tk.Text(xFrame)
        emailsText = tk.Text(xFrame)
        sendButton = tk.Button(self.root, text="Open Excel file", command=lambda:self.chooseFile(namesText, emailsText))
        sendButton.pack()

     

        namesText.grid(row=0, column =0, sticky= tk.W+tk.E)
        emailsText.grid(row=0, column =1, sticky= tk.W+tk.E)
        xFrame.pack()



    def chooseFile(self, names_text_widget, emails_text_widget): 
        filename = askopenfile()
        print("We are attempting to read the file: \n", filename)
        
        email_list = pd.read_excel(filename.name)
        names = email_list['name']
        change_text(names_text_widget, names)
        
        emails = email_list['email']
        change_text(emails_text_widget, emails)
#     print("The names we identified in the file are:\n", names)
#     print()
#     print("The emails we identified are:\n", emails)
#     print()

def change_text(text_widget, new_text):
    # Delete the current content from the beginning ('1.0') to the end ('end')
    text_widget.delete('1.0', 'end')
    # Insert the new text at the beginning ('1.0')
    text_widget.insert('1.0', new_text)

def main():
    root = tk.Tk()
    app = MyGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()

   
   
# def main():


#     B = Button(top, text ="Open Excel file", command = choose_file())
#     B.place(x=50,y=50)

#     print(B.command)
    
#     # send_emails(filename)

# # close the smtp server
#     print("Test run concluded")

# def choose_file():
#     filename = askopenfile()
#     print("We are attempting to read the file: \n", filename)
#     return filename

# def send_emails(filename):
#     outlook = win32.Dispatch('Outlook.Application')
#     email_list = pd.read_excel(filename.name)
    
#     names = email_list['name']
#     emails = email_list['email']

#     print("The names we identified in the file are:\n", names)
#     print()
#     print("The emails we identified are:\n", emails)
#     print()
 

#     for i in range(len(emails)):
 
#     # for every record get the name and the email addresses
#         mail = outlook.CreateItem(0)
#         mail.To = emails[i]
#         mail.Subject = 'Do not forget you have an unpaid invoice, dear ' + names[i]
#         mail.Body = 'MRAndom text 2, yey'
 
#     # # sending the email
#     mail.Send()


# if __name__ == '__main__':
#     main()