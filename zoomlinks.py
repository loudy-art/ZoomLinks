import pandas as pd
import smtplib
from email.mime.text import MIMEText
import tkinter as tk 
from tkinter.filedialog import askopenfilename
from tkinter import messagebox

#Gmail credentials
your_email = "example@gmail.com"
your_password = "yourpassword" #It is adviced using the app password that gmail can provide you so you don't have trouble sending e-mails through that address

#Reading the spreadsheet
root = tk.Tk()
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)
root.destroy()
email_list = pd.read_excel(filename, engine='openpyxl')

#Getting names and e-mails from file
names = email_list['NAME']
emails = email_list['MAIL']
link = input()

#Iterate through the records
for i in range(len(emails)):

    # For every record get the name and the email addresses
    name = names[i]
    email = emails[i]
    # Message to be emailed
    msg = MIMEText("Hi " + name + " this is the url: " + link + " have fun!") #Email body, you can use some html in here too but i'm lazy enough
    msg['Subject'] = 'Your class has arrived!' #Whatever subject you want
    
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #Use gmail with port
    session.starttls() #enable security
    session.login(your_email, your_password) #Login with mail_id and password
    message = msg.as_string()
    session.sendmail(your_email, [email], message)
    session.quit()
    print('Mail Sent')
    
