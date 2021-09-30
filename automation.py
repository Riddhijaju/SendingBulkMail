import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
data= pd.read_excel("mymailproject.xlsx")
print(type(data))
email_col=data.get("email")
list_of_email=list(email_col)
print(list_of_email)

try:
    server=sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("riddhi13jaju9@gmail.com","riddhi14201227jaju")
    from_="riddhi13jaju9@gmail.com"
    to_=list_of_email
    message=MIMEMultipart("alternative")
    message["Subject"]="my project to send bulk mails"
    message["from"]="riddhi13jaju9@gmail.com"

    html='''
    <html>
    <head>
    </head>
    <body>
    <h1>learn how to send email using python</h1>
    <p><h2>Getting Started</h2>
       Python comes with the built-in smtplib module for sending emails using the Simple Mail Transfer Protocol (SMTP). smtplib uses the RFC 821 protocol for SMTP. The examples in this tutorial will use the Gmail SMTP server to send emails, but the same principles apply to other email services. Although the majority of email providers use the same connection ports as the ones in this tutorial, you can run a quick Google search to confirm yours.

     To get started with this tutorial, set up a Gmail account for development, or set up an SMTP debugging server that discards emails you send and prints them to the command prompt instead. Both options are laid out for you below. A local SMTP debugging server can be useful for fixing any issues with email functionality and ensuring your email functions are bug-free before sending out any emails.

     Option 1: Setting up a Gmail Account for Development
     If you decide to use a Gmail account to send your emails, I highly recommend setting up a throwaway account for the development of your code. This is because you’ll have to adjust your Gmail account’s security settings to allow access from your Python code, and because there’s a chance you might accidentally expose your login details. Also, I found that the inbox of my testing account rapidly filled up with test emails, which is reason enough to set up a new Gmail account for development.

     A nice feature of Gmail is that you can use the + sign to add any modifiers to your email address, right before the @ sign. For example, mail sent to my+person1@gmail.com and my+person2@gmail.com will both arrive at my@gmail.com. When testing email functionality, you can use this to emulate multiple addresses that all point to the same inbox.

     To set up a Gmail address for testing your code, do the following:

     <h2>Create a new Google account.</h2>
     Turn Allow less secure apps to ON. Be aware that this makes it easier for others to gain access to your account.
     If you don’t want to lower the security settings of your Gmail account, check out Google’s documentation on how to gain access credentials for your Python script, using the OAuth2 authorization framework.

     ending emails to the specified address, it discards them and prints their content to the console. Running a local debugging server means it’s not necessary to deal with encryption of messages or use credentials to log in to an email server.
    </p>  
    </body>
    
    </html>
    
    
    
    
    '''
    part2=MIMEText(html,"html")




    message.attach(part2)

    server.sendmail(from_,to_,message.as_string())
    print("message has been sent to the emails.")

except Exception as e:
    print(e)
