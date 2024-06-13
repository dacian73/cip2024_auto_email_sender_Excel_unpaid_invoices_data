import win32com.client as win32

# Code for sending emails using Outlook
def send_emails(subjects, bodies, DATA):
        
        outlook = win32.Dispatch('Outlook.Application')
        for i in range(len(DATA)):
            # for every record create an email
            mail = outlook.CreateItem(0)
            mail.To = DATA[i]["email"]
            mail.Subject = subjects[i]
            mail.Body = bodies[i]

            print()
            print("For the item number", i, "we have the following email")
            print(DATA[i]["email"])
            print()
            print(subjects[i])
            print()
            print(bodies[i])
 
    # TODO sending the email
           # mail.Send()
    
# Code for preparing the emails
def on_send_emails_button_click(self, DATA, BODY_TEMPLATE):
        # Create lists for email subjects and bodies with the same number of elements as the emails list
        subjects = ["Unpaid Invoice Notification" for i in range(len(DATA))]
        bodies = ["" for i in range(len(DATA))]

        # Populate the lists with subjects and bodies as needed
        for i, client in enumerate(DATA):

            # Create an email body using the template and replacing the name and invoice data where needed
            body = ""
            for sequence in BODY_TEMPLATE:
                # We replace $name with the name of the client. We also cover the case when there is a comma, or something else attached to the end of the name
                if sequence[:5] == "$name":
                    remains = ""
                    if len(sequence)>5:
                        remains = sequence[5:]
                    body = body + client["name"] + remains
                elif sequence == "$invoices ":
                    displayable_invoices = ""
                    for invoice in client["invoices"]:
                        displayable_invoices = displayable_invoices + "\nInvoice number " + str(invoice["invoice_id"]) + " for " + str(invoice["sum"]) + "$" + " with due date " + str(invoice["date"].date())
                    body = body + displayable_invoices + " "
                else:
                    body = body+sequence
            
            bodies[i] = body
        send_emails(subjects, bodies, DATA)