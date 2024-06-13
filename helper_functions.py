import pickle
from tkinter import simpledialog
from tkinter.filedialog import asksaveasfile

from constants import DEFAULT_TEMPLATE_FILE

## TEMPLATES
# Function to load the default template
def load_default_template(file_path):
    try:
        with open(file_path, 'r') as file:
            return file.read()
    except FileNotFoundError:
        print(f"Template file {file_path} not found.")
        return ""

# Function to save a template
def save_template(BODY_TEMPLATE):
    file = asksaveasfile(initialfile="template")
    file.writelines(BODY_TEMPLATE)

# Function to save a template a default (overwrites the default template file)
def save_default_template(BODY_TEMPLATE):
    file = open(DEFAULT_TEMPLATE_FILE, 'w')
    file.writelines(BODY_TEMPLATE)

## USER DATA LISTS
# Save a list of users and invoice data
def save_list(DATA):
    filename = asksaveasfile(initialfile="list").name
    with open(filename, 'wb') as file:
        pickle.dump(DATA, file)

