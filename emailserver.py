# -*- coding: utf-8 -*-
import smtplib, ssl
import mimetypes
import numpy as np
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

import six
from pyfiglet import figlet_format
from email.mime.application import MIMEApplication

try:
    from termcolor import colored
except ImportError:
    colored = None

# database = pd.read_csv("Placement Database 2021 - Main Database.csv")
# database = database.iloc[1:100,::]
#  #                starting: ending , 169, 200
# database["HR Name"]= database["HR Name"].fillna("Hiring Manager")# "" --> hiring manager 
# details= dict(zip(database["Email"],database["HR Name"]))

# to_emails = database["Email"].values


login_details= {'smtp_server': 'smtp.gmail.com', 'smtp_protocol': 'tls', 'smtp_account': 'manav81101@gmail.com', 'smtp_password': 'upiydgfpnnvqbudy', 'example_no': 
'3', 'from_email': 'manav81101@gmail.com', 
'to_email': ["manavar81101@gmail.com"],
#"pranjalkukreja30@gmail.com","riyahimmatramka123@gmail.com","shreya190704@keshav.du.ac.in","sadiqansari605@gmail.com"],
# ["manav81101@gmail.com","pranjalkukreja30@gmail.com","riyahimmatramka123@gmail.com","shreya190704@keshav.du.ac.in",
#  "raghavrraghs@gmail.com ",
#  "haardikasethi@gmail.com",
#  "deeptijain9885@gmail.com",
# "sadiqansari605@gmail.com",] 
#"riyahimmatramka123@gmail.com",
# "shreya190704@keshav.du.ac.in",
# "raghavrraghs@gmail.com ",
# ",haardikasethi@gmail.com",
# "deeptijain9885@gmail.com",
# "sadiqansari605@gmail.com",
# "pranjalkukreja@gmail.com"], 
'subject': 'Invitation for Placement Fair 2020'}

details={"manavar81101@gmail.com":"Manav","pranjalkukreja30@gmail.com": "Pranjal","riyahimmatramka123@gmail.com": "Riya","shreya190704@keshav.du.ac.in":"Shreya","sadiqansari605@gmail.com":"Sadiq"}

PLAINT_TEXT_EMAIL = """
    Hi there,

    This message is sent from The Placement cell of Keshav Mahavidyalaya.

    Have a good day!
"""

def get_html_message(from_email, to_email, subject):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    # Plain-text version of content
    plain_text = """\
        Hi there,

        This message is sent from The Placement cell of Keshav Mahavidyalaya using Python
        

        Have a good day!
    """
    # html version of content
    html_content = """\
    <!DOCTYPE html>
    <html lang="en">
        <head>
            <style>
            /* .main-div{
                background-color:white;
            }
            body{
                background-color:white;
            }
            .div-inside{
                width: 50%;
            } */ 

            @media only screen and (max-width: 768px) {
                .main-div{
                    width:100%;
                }
                .div-inside{
                    width: 100%;
                }
                .img-fluid {
                 width: 100%;
                }
                .text-class{
                    font-size: 11px;
                }
            }
            @media only screen and (min-width: 768px) {
  /* For desktop: */
                .main-div{
                    width:50%;
                }
                .div-inside{
                    width: 100%;
                }
                .img-fluid {
                 width: 100%;
                }
                .text-class{
                    font-size: 17px;
                }
            }

            </style>
            <!-- <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css" integrity="sha384-TX8t27EcRE3e/ihU7zmQxVncDAy5uIKz4rEkgIXeMed4M0jlfIDPvg6uqKI2xXr2" crossorigin="anonymous">
            jQuery library--> 
            <!-- <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script> -->

            <!--Latest compiled and minified JavaScript--> 
            <!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script> -->


            <meta charset="UTF-8">
            <!-- <link rel="stylesheet" href="css/bootstrap.css"> -->
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Document</title>
        </head>
        <body>
            <div class="main-div">
                <div class=" div-inside">
                    <img class= "img-fluid"  src="https://i.ibb.co/WG52VhC/recruiters.jpg"   alt="Past Recruiters" border="0">
                    <p class= "text-class" style="margin:0; font-family:nunito;">Dear Manager</p>
                    <img class= "img-fluid" src="https://i.ibb.co/KzzT7XW/Greeting-from-Start-KMV-The-Placement-Cell-Of-Keshav-Mahavidyalaya-University-of-Delhi-We-would-like.png" alt="The Placement Cell of Keshav Mahavidyalaya"  border="0">
                </div>

            </div>

        </body>
    </html>
    """
    html_content=html_content.replace( "Manager",details[to_email],1)
    text_part = MIMEText(plain_text, 'plain')
    html_part = MIMEText(html_content, 'html')
    msg.attach(text_part)
    msg.attach(html_part)
    return msg

def get_attachment_message(from_email, to_email, subject):
    msg = get_html_message(from_email, to_email, subject)
    # Define MIMEImage part
    # Remember to change the file path
    file_path = './assets/KMV_PLACEMENT_BROCHURE.pdf'
    ctype, encoding = mimetypes.guess_type(file_path)
    maintype, subtype = ctype.split('/', 1)
    pdf = MIMEApplication(open(file_path, 'rb').read())
    pdf.add_header('Content-Disposition', 'attachment', filename='Placement Brochure.pdf')
    msg.attach(pdf)
    
    #print(msg)
    return msg

def cli_print(string, color, font="slant", figlet=False):
    if colored:
        if not figlet:
            six.print_(colored(string, color))
        else:
            six.print_(colored(figlet_format(
                string, font=font), color))
    else:
        six.print_(string)

def send_email(email_info):
    smtp_server = email_info.get('smtp_server', '')
    protocol = email_info.get('smtp_protocol', '')
    username = email_info.get('smtp_account', '')
    password = email_info.get('smtp_password', '')
    example_no = email_info.get('example_no', '')
    from_email = email_info.get('from_email', '')
    to_email = email_info.get('to_email', '')
    subject = email_info.get('subject', '')

    # Create a secure SSL context
    context = ssl.create_default_context()

    cli_print("********************************************", "green")

    try:
        if protocol == 'ssl':
            port = 465
            server = smtplib.SMTP_SSL(smtp_server, port, context=context)
            cli_print("Connecting to SMTP Server By Using SSL...", "green")
            server.login(username, password)
        elif protocol == 'tls':
            port = 587
            server = smtplib.SMTP(smtp_server, port)
            cli_print("Connecting to SMTP Server  By Using TLS...", "green")
            server.starttls(context=context) # Secure the connection with TLS
            server.login(username, password)
    except Exception as e:
        cli_print("Could not connect to SMTP server with exception: %s" % e, "red")
    else:
        cli_print("Sending your email...", "green")
        if example_no == '1':
            body = PLAINT_TEXT_EMAIL
            server.sendmail(from_email, to_email, body)
        elif example_no == '2':
            for x in to_email:
                msg = get_html_message(from_email, x, subject)
                server.send_message(msg)
        elif example_no == '3':
            for x in to_email:
                msg = get_attachment_message(from_email, x, subject)
                server.send_message(msg)
                cli_print("Your email was sent to " + x, "green")
    finally:
        server.quit()

def main():
    """
    Simple CLI for sending emails with Python
    """
    cli_print("Sending Email CLI", color="blue", figlet=True)
    cli_print("*** Welcome to Sending Email With Python ***", "green")

    email_info = login_details
    send_email(email_info)


if __name__ == "__main__":
    main()
