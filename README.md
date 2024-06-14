# Final Project for Code in Place 2024 - WORK IN PROGRESS

## Aim of the App
The app will allow users to send personalized emails to a large list of recipients. The emails will contain data about the unpaid invoices of each recipient.

## Libraries Used
- **Tkinter**: For the GUI
- **pandas**: For getting data from Excel files
- **win32com.client**: For sending emails with Outlook
- **pickle**: For writing and reading from file the object that contains client data

## Input
The user can choose an Excel file with columns named "client id", "name", "email", "invoice id", "sum", "due date". The file can contain multiple rows, including multiple rows with the same client, but with different invoices.

The user can also save data from the app in a file and load it back another time.

The app opens a default template from a file. The template can be edited and then saved by the user for later use.

## The GUI
 Please keep in mind that this was built by someone who just started learning python, in a short time, and it is a work in progress with the main goal being to learn.

 ### Starting the app

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/1.png?raw=true "At the start you can see the default template")

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/2.png?raw=true "The format of the excel file")

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/3.png?raw=true "After loading the data. You can see that users with multiple invoices only appear once, and their invoices appear in the last column.")

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/4.png?raw=true "Saving or loading a template or a list of invoices data")

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/5.png?raw=true "Small menu with a few options")

![Alt text](https://github.com/dacian73/cip2024_auto_email_sender_Excel_unpaid_invoices_data/blob/main/screenshots/7.png?raw=true "The end result: An email sent from the app.")
 

## How It Works
The user needs to choose a file from which the app will import data and store it in a global variable called data.

The structure of the data variable
data = [
    {
        "client_id" = client_id1
        "name" = name1, 
        "email" = email1, 
        "invoices" = [
            {"invoice_number": invoice_number1, "sum":sum1, "date":date1}, 
            {"invoice_number": invoice_number2, "sum":sum2, "date"=date2},
            ...
        ]
    },
    ...
]
Explanation:
 - data is a list of dictionaries
    - Each dictionary inside the data variable contains the keys:
        - "client_id"
        - "name"
        - "email"
        - "invoices" - list
            - The invoices list contains the details of each invoice from a client, the information being stored in a dictionary with the keys:
                - "invoice_id"
                - "sum"
                - "date"

 There is a predefined email body that is shown on a Text widget. The email body is editable. The user can write whathever he wants and he can use $name in place of the client name. The email body is stored in a global variable called body_template

 When pressing the send button, the app creates the emails and sends them.

 ### Description of the code
 #### main.py
    There are two main global variables: 
        DATA - which will store client id, name, email and invoices from an excel file or from a saved list
        BODY_TEMPLATE - which contains the current email template. When the app starts, the template is loaded from the default_template file
    We define the size, title and icon of the window.
    We create a menu with several options: exit, about and instructions
    We create a header
    We create a section for manipulating the email template. It contains a Text widget that can be edited and which trigers the function update_body_template each time it is modified by the user. The template is stored as a list in the variable BODY_TEMPLATE, and it is processed in order to be displayed in the Text widget in a readable format.
    We have three buttons for saving the template, replacing the default widget with the current one, or loading another template.
    The code for loading a new template can be found here and it also updates the Text widget.
    The code for saving templates can be found in the helper_functions.py

    We create 3 ListBox widgets, a button to load data from an excel, a button to save the data we are currently working with, and a button to load data that we previously saved.

    There is also a button to send the emails. The code to prepare and send the emails can be found in send_emails.py

 ### Icons
 The icons can be found in a folder called icons

 ### Constants.py
 Some of the constant values are stored here. More should be added.

 ### default_template
 The template that is opened when starting the app can be found here. The user can use the dedicated button to overwrite the default template.

 ### help_window.py
 Contains just some details about the app

 ### instructions_window.py
 Should contain instructions on how to use the app...

 ### send_emails.py
 Prepares the email bodies, by replacing $name from the template with the name of the user and the $invoices with invoice details.

 ## Further improvments
 There are a large number of potential improvements that could make this app more useful. Here are some of them:
 - making sure to handle most common errors
 - offer the option to use smtp instead of outlook
 - allow the user to add or delete clients from the imported lists
 - give the option to add pdf attachments with the invoices 
 - allow for custom email signature, that can be saved
 - make the number of columns dinamic, the user being able to decide what information he wants to use or see
 - give the user the option to write the names of the column headers he is interested in (instead of forcing him to use "name", "email", "invoice id" etc.)
 - using a better logging system, instead of the many chaotic prints
