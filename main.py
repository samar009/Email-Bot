import os
import pandas as pd
import smtplib
from email.message import EmailMessage

#Set your name, Email ID and Password here.
your_name = "<name>"
your_email = "<email id>"
your_password = "<password>"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(your_email, your_password)

e= pd.read_excel("Email.xlsx")
email =e['Emails'].values

msg = EmailMessage()
msg['Subject'] = 'subject'
msg['From'] = 'email address'
msg['To'] = email
#msg.set_content('Holiday List, HAPPY HOLIDAYS!!')
msg.add_alternative("""<!DOCTYPE html>
<html>
    <body>
        <h1 style="color:SlateGray;">HAPPY HOLIDAYS!!</h1>
    </body>
</html>
""",subtype='html')

with open('Holidays.xlsx','rb') as f:
    file_data = f.read()
msg.add_attachment(file_data,maintype='application',subtype='octet-stream',filename='Holidays.xlsx')
try:
    server.send_message(msg)
    print('Email  to {} successfully sent!\n\n'.format(email))
except Exception as e:
    print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

server.close()