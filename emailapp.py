from tkinter import *
import tkinter as tk
startNumber = 0
endingNumber = 0
counter = 0 

def random(style, body):
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
    html_content= html_content.replace( "Manager","Manav",1)
    THTML.insert(END, html_content)


def ButtonClick(*args):
    startNumber = txtfldstarting.get()
    endingNumber = txtfldending.get()
    txtem = txtfldem.get()
    txtpass = txtfldpass.get()
    sheetlink= txtfldshtlink.get()
    subject = txtfldSubject.get()
    style = Tstyle.get('1.0', 'end-1c')
    body = Tbody.get('1.0', 'end-1c')
    #lb.insert(END, style,body)
    THTML.delete('1.0', END)
    random(style, body)
    


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

btn=Button(window, text="Send Mails", fg='blue')
btn.place(x=435, y=700)
labelwarning = Label(window, text ="Please click the button only once" , fg='black', font=("karla", 16))
labelwarning.place(x=300, y= 750)

btn.bind('<Button-1>', ButtonClick)


lb=Listbox(window,width= 70, selectmode='multiple')
lb.place(x=250, y=800)


labelhead = Label(window,text="Enter the complete style tag here", fg='black', font=("karla", 16))
labelhead.place(x=75 , y= 350)
Tstyle = tk.Text(window, height=10, width=50)
Tstyle.place(x = 50, y= 400)
Tstyle.insert(tk.END, "Enter the style tag here")

labelhead = Label(window,text="Enter the complete body here", fg='black', font=("karla", 16))
labelhead.place(x=575 , y= 350)
Tbody = tk.Text(window, height=10, width=50)
Tbody.place(x = 550, y= 400)
Tbody.insert(tk.END, "Enter the body here")

THTML = tk.Text(window, height=10, width=50)
THTML.place(x = 100, y= 100)

window.title('Email Server Application')
window.geometry("1000x1000")
window.mainloop()