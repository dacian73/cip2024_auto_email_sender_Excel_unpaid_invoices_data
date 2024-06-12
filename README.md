# Final Project for Code in Place 2024 - WORK IN PROGRESS

## Aim of the App
The app will allow users to send personalized emails to a large list of recipients. The emails will contain data about the unpaid invoices of each recipient.

## Libraries Used
- **Tkinter**: For the GUI
- **pandas**: For getting data from Excel files
- **win32com.client**: For sending emails with Outlook

## Input
The user can choose an Excel file with columns named "name", "email", and "sum". The file can contain multiple rows, including multiple rows with the same "name" and "email", but with different invoice sums and dates.

## The GUI
*...GUI description goes here...*

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
        - "client_id" - String
        - "name"  - String
        - "email" - String
        - "invoices" - list
            - The invoices list contains the details of each invoice from a client, the information being stored in a dictionary with the keys:
                - "invoice_id" - String
                - "sum" - String
                - "date" - String



 There is a predefined email body that is shown on a Text widget. The email body is editable. The user can write whathever he wants and he can use $name in place of the client name. The email body is stored in a global variable called body_template

 When pressing the send button, the app creates the emails and sends them.

 ## Further improvments
 There are a large number of potential improvements that could make this app more useful. Here are some of them:
 - making sure to handle most common errors
 - offer the option to use smtp instead of outlook
 - allow the user to add or delete clients from the imported lists
 - give the option to add pdf attachments with the invoices 
 - allow for custom email signature, that can be saved
 - make the number of columns dinamic, the user being able to decide what information he wants to use or see
 - give the user the option to write the names of the column headers he is interested in (instead of forcing him to use "name", "email", "invoice id" etc.)

