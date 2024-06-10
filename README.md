# Final Project for Code in Place 2024 - WORK IN PROGRESS

## Aim of the App
The app will allow users to send personalized emails to a large list of recipients. The emails will contain data about the unpaid invoices of each recipient.

## Libraries Used
- **Tkinter**: For the GUI
- **openpyxl**: For opening Excel files

## Input
The user can choose an Excel file with columns named "name", "email", and "sum". The file can contain multiple rows, including multiple rows with the same "name" and "email", but with different invoice sums and dates.

## The GUI
*...GUI description goes here...*

## How It Works
The user needs to choose a file from which the app will import data and store it in a global variable called data.

The structure of the data variable
data = [
    {
        "name" = name, 
        "email" = email, 
        "invoices" = [
            {"sum":sum1, "date":date1}, {"sum":sum2, "date"=date2},
            ...
        ]
    },
    ...
]
Explanation:
 - data is a list of dictionaries
    - Each dictionary inside the data variable contains the keys:
        - "name"  - String value
        - "email" - String value
        - "invoices" - list
            - The invoices list contains the details of each invoice from a client, the information being stored in a dictionary with the keys:
                - "sum" - Stirng value
                - "date" - String value



 There is a predefined email body that is shown on a Text widget. The email body is editable. The user can write whathever he wants and he can use $name in place of the client name. The email body is stored in a global variable called body_template

 When pressing the send button, the app creates the emails and sends them.