from tkinter import simpledialog
import win32com.client as win32
from tkinter.filedialog import askopenfile, askopenfilename, asksaveasfile
import tkinter as tk
import pandas as pd
import pickle

# Imports from project modules
from constants import *
from help_window import helpPage
from helper_functions import load_default_template, save_default_template, save_list, save_template
from instructions_window import instructionsPage
from send_emails import *

DATA = []
global BODY_TEMPLATE

# The default template for the email is loaded from a file
BODY_TEMPLATE = load_default_template(DEFAULT_TEMPLATE_FILE)

# The user interface
class MyGui:
    def __init__(self, master):
        self.root = master
        self.setup_gui()

    def setup_gui(self):
        # Tkinter window details
        self.root.geometry(WINDOW_SIZE)
        self.root.title(WINDOW_TITLE)
        self.root.iconbitmap(ICON_PATH)
        self.create_menu()
        self.create_header()
        self.create_email_template_section()
        self.create_data_loading_section()
    
    def create_menu(self):
        # Menu Bar
        self.menubar = tk.Menu(self.root)
        self.filemenu=tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Exit", command=self.root.destroy)
        self.menubar.add_cascade(menu=self.filemenu, label="File")
        self.helpmenu=tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="About", command=lambda:helpPage(self))
        self.helpmenu.add_command(label="Instructions", command=lambda:instructionsPage(self))
        self.menubar.add_cascade(menu=self.helpmenu, label="Help")
        self.root.config(menu=self.menubar)

    def create_header(self):
        # Header
        header = tk.Label(self.root, text="Bulk email sender", font=("Arial", 18))
        header.pack(padx=20, pady=20)

    def create_email_template_section(self):
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
        for sequence in BODY_TEMPLATE:
                if sequence == "$name":
                    displayable_body_template = displayable_body_template + "$name"
                elif sequence == "$invoices":
                    displayable_body_template = displayable_body_template + "$invoices"
                else:
                    displayable_body_template = displayable_body_template+sequence
        # Using displayable_body_template as text for the templateText widget
        templateText.insert("end-1c", displayable_body_template)
        templateText.pack(padx=10, pady=10)
        # Function to update the body_template variable with the content from the templateText widget
        def update_body_template(event):
            updated_text = templateText.get("1.0", tk.END)
            global BODY_TEMPLATE
            BODY_TEMPLATE = [""]
            for sequence in updated_text.split():
                BODY_TEMPLATE.append(sequence + " ")
            event.widget.edit_modified(False)
        # Binding templateText, so that body_template gets updated each time the text is modified
        templateText.bind('<<Modified>>', update_body_template)

        # Function for loading a template and updating the templateText widget
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

        saveButton = tk.Button(buttonFrame, text="Save Template", command=lambda:save_template(BODY_TEMPLATE))
        saveAsDefaultButton = tk.Button(buttonFrame, text="Save as Default Template", command=lambda:save_default_template(BODY_TEMPLATE))
        loadButton = tk.Button(buttonFrame, text="Load Template", command=load_template)
        saveButton.grid(row=0, column =0, sticky= tk.W+tk.E)
        saveAsDefaultButton.grid(row=0, column =1, sticky= tk.W+tk.E)
        loadButton.grid(row=0, column =2, sticky= tk.W+tk.E)

        buttonFrame.pack()

    def create_data_loading_section(self):
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

        saveListButton = tk.Button(saveListFrame, text="Save List", command=lambda:save_list(DATA))
        loadListButton = tk.Button(saveListFrame, text="Load List", command=lambda:self.load_data_from_file(namesListBox,emailsListBox,sumsListBox))
        saveListButton.grid(row=0, column =0, sticky= tk.W+tk.E)
        loadListButton.grid(row=0, column =1, sticky= tk.W+tk.E)
        saveListFrame.pack()

        # Button to send emails
        sendEmailsButton = tk.Button(self.root, text="Send Emails", command=lambda:on_send_emails_button_click(self, DATA, BODY_TEMPLATE))
        sendEmailsButton.pack()

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

        global DATA

        for i in range(len(names)):
            if names[i] in names_copy:
                print("We identified another invoice with client id =", client_ids[i])
                index = next((index for (index, d) in enumerate(DATA) if d["name"] == names[i]), None)
                print("the index is =", index)
                print("data[index]=", DATA[index])
                DATA[index].get('invoices').append({"invoice_id": invoice_ids[i],"sum":sums[i], "date": dates[i]})
            else:
                names_copy.append(names[i])
                DATA.append({"client_id": client_ids[i],"name":names[i], "email":emails[i], "invoices":[{"invoice_id": invoice_ids[i],"sum":sums[i], "date": dates[i]}]})

        invoices = [ value["invoices"] for value in DATA ]
        names = [ value["name"] for value in DATA ]
        emails = [ value["email"] for value in DATA ]

        print("The names we identified in the file are:\n", names)
        print()
        change_list(names_listbox, names)
        

        print("The emails we identified are:\n", emails)
        print()
        change_list(emails_listbox, emails)

        print("The invoices we identified are:\n", invoices)
        print()
        displayabe_invoices = []
        for all_user_invoices in invoices:
            one_user_invoices = ""
            for invoice in all_user_invoices:
                one_user_invoices = one_user_invoices + "Invoice number " + str(invoice["invoice_id"]) + " for " + str(invoice["sum"]) + ". "
            displayabe_invoices.append(one_user_invoices)
        change_list(sums_listbox, displayabe_invoices)

    # Function for loading a Data from a file
    def load_data_from_file(self, names_listbox, emails_listbox, invoices_listbox):
            global DATA
            filename = askopenfilename()
            if filename:
                try:
                    with open(filename, 'rb') as file:
                        DATA = pickle.load(file)
                        invoices = [ value["invoices"] for value in DATA ]
                        names = [ value["name"] for value in DATA ]
                        emails = [ value["email"] for value in DATA ]
                        change_list(names_listbox, names)
                        change_list(emails_listbox, emails)
                        displayabe_invoices = []
                        for all_user_invoices in invoices:
                            one_user_invoices = ""
                            for invoice in all_user_invoices:
                                one_user_invoices = one_user_invoices + "Invoice number " + str(invoice["invoice_id"]) + " for " + str(invoice["sum"]) + ". "
                            displayabe_invoices.append(one_user_invoices)
                        change_list(invoices_listbox, displayabe_invoices)
                except FileNotFoundError:
                    print(f"File '{filename}' not found.")    

# Update the content of ListBox widgets
def change_list(list_widget, new_list):
    # Delete the current content for the ListBox widget
    list_widget.delete("0", "end")
    # Insert the new list for the ListBox widget
    for list in new_list:
        list_widget.insert('0', list)


def main():
    root = tk.Tk()
    app = MyGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()