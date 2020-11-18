# -*- coding: utf-8 -*-
import smtplib, ssl
import mimetypes
import numpy as np
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tkinter import *
import tkinter as tk
import six
from pyfiglet import figlet_format
from email.mime.application import MIMEApplication

try:
    from termcolor import colored
except ImportError:
    colored = None


styletagsfinal = """<style>
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

            </style>"""

bodytagsfinal = """<body>
            <div class="main-div">
                <div class=" div-inside">
                    <img class= "img-fluid"  src="https://i.ibb.co/WG52VhC/recruiters.jpg"   alt="Past Recruiters" border="0">
                    <p class= "text-class" style="margin:0; font-family:nunito;">Dear Manager</p>
                    <img class= "img-fluid" src="https://i.ibb.co/KzzT7XW/Greeting-from-Start-KMV-The-Placement-Cell-Of-Keshav-Mahavidyalaya-University-of-Delhi-We-would-like.png" alt="The Placement Cell of Keshav Mahavidyalaya"  border="0">
                </div>

            </div>

        </body>
"""


login_details = dict()
to_emails = np.zeros(3)
details = dict()

PLAINT_TEXT_EMAIL = """
    Hi there,

    This message is sent from The Placement cell of Keshav Mahavidyalaya.

    Have a good day!
"""

def emails_to_be_sent(snum, enum, mailaddr, passwd, sheetlink, subject):
    try:
        database = pd.read_csv(sheetlink)
        lb.insert(END,"Database Loaded successfully")
        try:
            snum= int(snum)
            enum= int(enum)
            database = database.iloc[snum-1:enum,::]
        #                starting: ending , 169, 200
            database["HR Name"]= database["HR Name"].fillna("Hiring Manager")# "" --> hiring manager 
            details= dict(zip(database["Email"],database["HR Name"]))
            to_emails = database["Email"].values
            login_details= {'smtp_server': 'smtp.gmail.com', 'smtp_protocol': 'tls', 'smtp_account': mailaddr, 'smtp_password': passwd, 'example_no': '3', 'from_email': mailaddr, 'to_email': to_emails, 'subject': subject}

        except:
            lb.insert(END,"starting and ending numbers are not correct ")
    except:
        lb.insert(END, "Please enter valid URL ")
    return (login_details, details)


def get_html_message(from_email, to_email, subject, detailinfo, style, body):
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
                styletags
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
            bodytags
        </html>
        
    """
    html_content = html_content.replace("styletags", style,1 )
    html_content = html_content.replace("bodytags", body,1 )
    html_content=html_content.replace( "Manager",detailinfo[to_email],1)
    text_part = MIMEText(plain_text, 'plain')
    html_part = MIMEText(html_content, 'html')
    msg.attach(text_part)
    msg.attach(html_part)
    return msg

def get_attachment_message(from_email, to_email, subject, detailinfo, style, body):
    msg = get_html_message(from_email, to_email, subject, detailinfo, style, body)
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

def send_email(email_info, detailinfo, style,body):
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
            lb.insert(tk.END,"Your have logged in with your account successfully", username)
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
                msg = get_attachment_message(from_email, x, subject, detailinfo,style,body)
                server.send_message(msg)
                cli_print("Your email was sent to " + x, "green")
                lb.insert(tk.END,"Your email was sent to " + x + " Successfully")
    finally:
        server.quit()

def main(start ,end , mailaddr, passwd, sheetlink, subject, style, body):
    """
    Simple CLI for sending emails with Python
    """
    cli_print("Sending Email CLI", color="blue", figlet=True)
    cli_print("*** Welcome to Sending Email With Python ***", "green")
    loginreturn = emails_to_be_sent(start ,end, mailaddr,passwd, sheetlink, subject)
    email_info, detailinfo = loginreturn
    send_email(email_info, detailinfo, style,body)

#main()
if __name__ == "__main__":
    # pass
    startNumber = 0
    endingNumber = 0 
    def ButtonClick(*args):
        startNumber = txtfldstarting.get()
        endingNumber = txtfldending.get()
        txtem = txtfldem.get()
        txtpass = txtfldpass.get()
        sheetlink= txtfldshtlink.get()
        subject = txtfldSubject.get()
        style = Tstyle.get('1.0', 'end-1c')
        body = Tbody.get('1.0', 'end-1c')
        # lb.insert(END, style,body)
        main(startNumber,endingNumber, txtem,txtpass, sheetlink, subject, style, body)
        

    
    window=Tk()
    # add widgets here
    labelSubject =Label(window, text="Enter the Subject here", fg='black', font=("karla", 16))
    labelSubject.place(x=350, y=150)
    
    txtfldSubject =Entry(window, text="This is Subject", bg='white',fg='black', bd=5)
    txtfldSubject.place(x=400, y=200)


    labelsheetlink =Label(window, text="Enter the sheet link here", fg='black', font=("karla", 16))
    labelsheetlink.place(x=350, y=50)
    
    txtfldshtlink =Entry(window, text="This is sheetlink", bg='white',fg='black', bd=5)
    txtfldshtlink.place(x=400, y=100)


    labelemail =Label(window, text="Enter the email here", fg='black', font=("karla", 16))
    labelemail.place(x=60, y=50)

    txtfldem =Entry(window, text="This is email", bg='white',fg='black', bd=5)
    txtfldem.place(x=90, y=100)

    lblpass=Label(window, text="Enter the Password", fg='black', font=("karla", 16))
    lblpass.place(x=60, y=150)
    txtfldpass=Entry(window, text="This is Password", bg='white',fg='black', bd=5)
    txtfldpass.place(x=90, y=200)



    labelstarting =Label(window, text="Enter the Starting Number", fg='black', font=("karla", 16))
    labelstarting.place(x=690, y=50)
    txtfldstarting =Entry(window, text="This is Starting Number", bg='white',fg='black', bd=5)
    txtfldstarting.place(x=750, y=100)

    lblending=Label(window, text="Enter the Ending Number", fg='black', font=("karla", 16))
    lblending.place(x=690, y=150)
    txtfldending=Entry(window, text="This is Ending Number", bg='white',fg='black', bd=5)
    txtfldending.place(x=750, y=200)

    btn=Button(window, text="Send Mails", fg='black')
    btn.place(x=435, y=700)
    labelwarning = Label(window, text ="Please click the button only once" , fg='black', font=("karla", 16))
    labelwarning.place(x=300, y= 750)

    btn.bind('<Button-1>', ButtonClick)


    lb=tk.Text(window,width= 70)
    lb.place(x=250, y=800)

    labelhead = Label(window,text="Enter the complete style tag here", fg='black', font=("karla", 16))
    labelhead.place(x=75 , y= 350)
    Tstyle = tk.Text(window, height=10, width=50)
    Tstyle.place(x = 50, y= 400)
    Tstyle.insert(tk.END, styletagsfinal)

    labelhead = Label(window,text="Enter the complete body here", fg='black', font=("karla", 16))
    labelhead.place(x=600 , y= 350)
    Tbody = tk.Text(window, height=10, width=50)
    Tbody.place(x = 550, y= 400)
    Tbody.insert(tk.END, bodytagsfinal)



    window.title('Email Server Application')
    window.geometry("1000x1000")
    window.mainloop()






