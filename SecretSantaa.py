import pandas as pd
import random
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


data = pd.read_excel('NAME OF YOUR FILE.xlsx')



#assuming 'Name' column contains the names
names = list(data['Name'])
random.shuffle(names)
assignments = {names[i]: names[(i + 1) % len(names)] for i in range(len(names))}

#gmail
email_address = 'YOUR_EMAIL'
email_password = 'YOUR_PASSWORD'

#SMTP server setup
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email_address, email_password)

#Sending emails
for sender, recipient in assignments.items():
    #assuming 'GoogleDrive' column contains the Google Drive link 
    #i used a Google Form for Secret Santa questions then saved each response individually and uploaded to Google Drive
    recipient_drive_link = data.loc[data['Name'] == recipient, 'GoogleDrive'].values[0]

    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = data.loc[data['Name'] == sender, 'Email'].values[0]
    msg['Subject'] = 'Secret Santa Assignment'

    #edit the body to your rules
    body = f"Hello {sender},\n\nYou are the Secret Santa for {recipient}!\n"
    body += f"The price limit is $30 you can use {recipient}'s Secret Santa form here: {recipient_drive_link}\n"
    body += f"Merry Christmas!"
    msg.attach(MIMEText(body, 'plain'))

    server.sendmail(email_address, msg['To'], msg.as_string())

server.quit()


