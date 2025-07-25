import smtplib
import pandas as pd
from getpass import getpass





# Email account credentials
sender_email = "c3@cylab.be"
password = getpass("Enter your email password: ")


# Create the email
#subject = "Test send mail using python"
#body = "Hello yemen,  I made it ;) "


smtp_server = "mail.cylab.be"
smtp_port = 587

excel_file = "/home/ioanna/projects/pythonScriptingForExcel/testexcel.xlsx"


try:

      # Path to your Excel file
    
    data = pd.read_excel(excel_file)
    # Connect to the mail server
    with smtplib.SMTP(smtp_server, smtp_port) as server:

        server.set_debuglevel(1)

        server.starttls()  # Secure the connection if using TLS
        server.login(sender_email, password)  # Login


        for index, row in data.iterrows():
            
            # Customize your message
            recipient_email = row['Email address']  # Replace 'Name' with the actual column name for recipient names

            server.sendmail(sender_email, recipient_email, 'Subject: C3 Selection Results\n\nCongratulations!!! You have been selected to take part to the C3 workshop!')  # Send email
            
            print(f"Email to {recipient_email} sent successfully!")

except Exception as e:
    print(f"Failed to send email: {e}")


#server.quit()  # Logout of the email server

