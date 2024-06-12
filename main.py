from tkinter import simpledialog
import win32com.client as win32
from tkinter.filedialog import askopenfile, askopenfilename, asksaveasfile
import tkinter as tk
import pandas as pd

data = []
global body_template

# The default template for the email is loaded from a file
file = open('default_template', 'r')
body_template = file.read()

# The user interface
class MyGui:
    def __init__(self, master):
        
        # Tkinter window details
        self.root = master
        self.root.geometry("600x600")
        self.root.title("Email Sender app - CIP2024 project")
        self.root.iconbitmap("icon.ico")

        # Menu Bar
        self.menubar = tk.Menu(self.root)
        self.filemenu=tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Exit", command=self.root.destroy)
        self.menubar.add_cascade(menu=self.filemenu, label="File")
        self.helpmenu=tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="About", command=self.helpPage)
        self.menubar.add_cascade(menu=self.helpmenu, label="Help")
        self.root.config(menu=self.menubar)

        # Header
        header = tk.Label(self.root, text="Bulk email sender", font=("Arial", 18))
        header.pack(padx=20, pady=20)
        # Label: Email Template
        templateLabel = tk.Label(self.root, text="Email Template", font=("Arial", 12))
        templateLabel.pack()
        # Email template instructions
        editTemplateLabel = tk.Label(self.root, text="You can edit the template bellow. you need to write $name instead of the client name.", font=("Arial", 8))
        editTemplateLabel.pack()
        # Text widget with the email template
        templateText = tk.Text(self.root, font=("Arial", 10), height=10)
        # Transforming the list body_template into a more readable string called displayable_body_template
        displayable_body_template = ""
        for sequence in body_template:
                if sequence == "$name":
                    displayable_body_template = displayable_body_template + "$name"
                elif sequence == "$invoices":
                    displayable_body_template = displayable_body_template + "$invoices"
                else:
                    displayable_body_template = displayable_body_template+sequence
        # Using displayable_body_template as text for the templateText widget
        templateText.insert("end-1c", displayable_body_template)
        templateText.pack(padx=10, pady=10)
        # Change the body template when the text in the templateText widget changes
        def update_body_template(event):
            updated_text = templateText.get("1.0", tk.END)
            global body_template
            body_template = [""]
            for sequence in updated_text.split():
                body_template.append(sequence + " ")
            event.widget.edit_modified(False)
        # Binding templateText, so that body_template gets updated each time it is modified
        templateText.bind('<<Modified>>', update_body_template)

        # Loading a template
        def load_template():
            filename = askopenfilename()
            if filename:
                try:
                    with open(filename, 'r') as f:
                        data = f.read()
                        templateText.delete("1.0", "end")  # Clear existing content
                        templateText.insert("end", data)  # Insert new data
                except FileNotFoundError:
                    print(f"File '{filename}' not found.")
        
        # Buttons for saving or loading templates
        buttonFrame = tk.Frame(self.root)
        buttonFrame.columnconfigure(0, weight=1)
        buttonFrame.columnconfigure(1,weight=1)
        buttonFrame.columnconfigure(2,weight=1)

        saveButton = tk.Button(buttonFrame, text="Save Template", command=self.save_template)
        saveAsDefaultButton = tk.Button(buttonFrame, text="Save as Default Template", command=self.save_default_template)
        loadButton = tk.Button(buttonFrame, text="Load Template", command=load_template)
        saveButton.grid(row=0, column =0, sticky= tk.W+tk.E)
        saveAsDefaultButton.grid(row=0, column =1, sticky= tk.W+tk.E)
        loadButton.grid(row=0, column =2, sticky= tk.W+tk.E)

        buttonFrame.pack()

        # Button to load data from excel file
        openFileButton = tk.Button(self.root, text="Open Excel file", command=lambda:self.chooseFile(namesListBox, emailsListBox, sumsListBox))
        openFileButton.pack()

        # ListBox for data about clients and their invoices
        dataFrame = tk.Frame(self.root)
        dataFrame.columnconfigure(0, weight=1)
        dataFrame.columnconfigure(1,weight=1)
        dataFrame.columnconfigure(2,weight=1)

        namesListBox = tk.Listbox(dataFrame)
        emailsListBox = tk.Listbox(dataFrame)
        sumsListBox = tk.Listbox(dataFrame)

        namesListBox.grid(row=0, column =0, sticky= tk.W+tk.E)
        emailsListBox.grid(row=0, column =1, sticky= tk.W+tk.E)
        sumsListBox.grid(row=0, column =2, sticky= tk.W+tk.E)
        dataFrame.pack()

        # Buttons to save or load list
        saveListFrame = tk.Frame(self.root)
        saveListFrame.columnconfigure(0, weight=1)
        saveListFrame.columnconfigure(1, weight=1)

        saveListButton = tk.Button(saveListFrame, text="Save List", command=self.save_list)
        loadListButton = tk.Button(saveListFrame, text="Load List", command=self.load_list)
        saveListButton.grid(row=0, column =0, sticky= tk.W+tk.E)
        loadListButton.grid(row=0, column =1, sticky= tk.W+tk.E)
        saveListFrame.pack()



        # Button to send emails
        sendEmailsButton = tk.Button(self.root, text="Send Emails", command=self.on_send_emails_button_click)
        sendEmailsButton.pack()

    def save_template(self):
        file = asksaveasfile(initialfile="template")
        displayable_body_template = ""
        for sequence in body_template:
                if sequence == "$name":
                    displayable_body_template = displayable_body_template + "$name"
                elif sequence == "$invoices":
                    displayable_body_template = displayable_body_template + "$invoices"
                else:
                    displayable_body_template = displayable_body_template+sequence
        file.write(str(displayable_body_template))
    
    def save_default_template(self):
        file = open("default_template", 'w')
        file.writelines(body_template)

    def save_list(self):
        filename = simpledialog.askstring(title="Saving...",
                                  prompt="Write a name for the file you want to save")
        file = open(filename, 'w')
        file.writelines(data)

    def load_list(self):
        global data
        file = open('default_template', 'r')
        data = file.read()
        
    def helpPage(self):
        # Create a tkinter window for "About" info
        win=tk.Tk()
        win.geometry("600x400")
        win.title("About this app")
        win.iconbitmap("about_icon.ico")

        label = tk.Label(win, text= "About the app! ",font=('Arial bold', 18)).pack(pady=20)
        label = tk.Label(win, text= "This app was created by dacian73 for the Code in Place 2024 final project.\nThe sourcecode is available at https://github.com/dacian73 ",font=('Arial', 12)).pack(pady=20)
        #Make the window jump above all
        win.attributes('-topmost',True)
        win.mainloop()
    


    # Let the user upload an Excel file with columns for name, email
    def chooseFile(self, names_listbox, emails_listbox, sums_listbox): 
        filename = askopenfile()
        print("We are attempting to read the file: \n", filename)
        print()

        
        input_from_file = pd.read_excel(filename.name)
        print("Email list = ", input_from_file)

        client_ids = input_from_file['client id']

        invoice_ids = input_from_file['invoice id']

        sums = input_from_file['sum']

        dates = input_from_file['due date']

        names = input_from_file['name'].to_list()

        emails = input_from_file['email'].to_list()
        names_copy = []

        global data

        for i in range(len(names)):
            if names[i] in names_copy:
                print("We identified another invoice with client id =", client_ids[i])
                index = next((index for (index, d) in enumerate(data) if d["name"] == names[i]), None)
                print("the index is =", index)
                print("data[index]=", data[index])
                data[index].get('invoices').append({"invoice_id": invoice_ids[i],"sum":sums[i], "date": dates[i]})
            else:
                names_copy.append(names[i])
                data.append({"client_id": client_ids[i],"name":names[i], "email":emails[i], "invoices":[{"invoice_id": invoice_ids[i],"sum":sums[i], "date": dates[i]}]})

        invoices = [ value["invoices"] for value in data ]
        names = [ value["name"] for value in data ]
        emails = [ value["email"] for value in data ]

        print("The names we identified in the file are:\n", names)
        print()
        change_text(names_listbox, names)
        

        print("The emails we identified are:\n", emails)
        print()
        change_text(emails_listbox, emails)

        print("The invoices we identified are:\n", invoices)
        print()
        displayabe_invoices = []
        for all_user_invoices in invoices:
            one_user_invoices = ""
            for invoice in all_user_invoices:
                one_user_invoices = one_user_invoices + "Invoice number " + str(invoice["invoice_id"]) + " for " + str(invoice["sum"]) + ". "
            displayabe_invoices.append(one_user_invoices)
        change_text(sums_listbox, displayabe_invoices)

        

    def send_emails(self, subjects, bodies):
        
        outlook = win32.Dispatch('Outlook.Application')
        for i in range(len(data)):
            # for every record create an email
            mail = outlook.CreateItem(0)
            mail.To = data[i]["email"]
            mail.Subject = subjects[i]
            mail.Body = bodies[i]

            print()
            print("For the item number", i, "we have the following email")
            print(data[i]["email"])
            print()
            print(subjects[i])
            print()
            print(bodies[i])
 
    # TODO sending the email
           # mail.Send()
    
    def on_send_emails_button_click(self):
        # Create lists for email subjects and bodies with the same number of elements as the emails list
        subjects = ["Unpaid Invoice Notification" for i in range(len(data))]
        bodies = ["" for i in range(len(data))]

        # Populate the lists with subjects and bodies as needed
        for i, client in enumerate(data):

            # Create an email body using the template and replacing the name and invoice data where needed
            body = ""
            for sequence in body_template:
                if sequence == "$name ":
                    body = body + client["name"] + " "
                elif sequence == "$invoices ":
                    displayable_invoices = ""
                    for invoice in client["invoices"]:
                        displayable_invoices = displayable_invoices + "\nInvoice number " + str(invoice["invoice_id"]) + " for " + str(invoice["sum"]) + "$" + " with due date " + str(invoice["date"].date())
                    body = body + displayable_invoices + " "
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


def main():
    root = tk.Tk()
    app = MyGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()