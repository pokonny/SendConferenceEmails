import tkinter as tk
from tkinter import *
import win32com.client as win32
from icalendar import Calendar, Event, vCalAddress, vText
import datetime


# reads calendar invite, creates email_template, fills out info, then displays immage
def read_invite(file_name, email_template):
    # asks for ics file location opens ics file
    cal = Calendar()
    # Some properties are required to be compliant
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')
    conf_email_temp = outlook.CreateItemFromTemplate(email_template)
    meetorg = ''
    meet_location = ''
    formatted_date = ''
    formatted_time = ''
    # Finds the file and reads what kind it is
    msg = r'C:\Users\pokonny\Downloads\calendar(' + file_name + ').ics'
    e = open(msg, 'rb')
    ecal = cal.from_ical(e.read())

    # Walks through the file and adds it to the email
    for component in ecal.walk():
        if component.name == 'VEVENT':
            meetorg = component.get('organizer').replace('mailto:', '')
            meet_location = component.get('location')

            # Checks to see if the file contains a ; for when microsoft teams is included
            if ';' in meet_location:
                meet_locationa = meet_location.split(';')
                meet_location = meet_locationa[1]

            # Grabs meeting date and time then formats it in MM/DD/YYYY AND XX:YY AM/PM
            conf_email_temp.To = meetorg
            meet_date = component.decoded('dtstart')
            formatted_date = datetime.date.strftime(meet_date, '%m/%d/%Y ')
            formatted_time = datetime.date.strftime(meet_date, '%H:%M %p')
            spiel = meet_location + ' on ' + formatted_date + 'at ' + formatted_time + '. '
            conf_email_temp.Subject = 'Do you need Tech Assistance at ' + spiel
            conf_email_temp.Body += ' ' + spiel

    e.close()
    conf_email_temp.Display()


# prompts for the file location
def conference_email(email_template):
    file_nume = Tk()
    label = Label(file_nume, text="File location")
    label.pack(side=LEFT)
    file_name = Entry(file_nume, bd=5)
    file_name.pack(side=LEFT)
    b = Button(file_nume, text="Submit",
               command=lambda: [read_invite(file_name.get(), email_template), close(file_nume)])
    b.pack()
    file_nume.mainloop()


# reads user input, adds subject, inputs user input and labels, then displays message
def as400write(message, textbox):
    user = textbox.get("1.0", 'end').split('\n')
    message.Subject = 'Your AS400 access has been granted.'
    # removes recipient and cc emails from array
    message.To = user.pop(0)
    message.CC = user.pop(0)
    # Adds a new line before writing
    message.Body += "\n"
    c = ["Username: ", "Password: ", "AS400 Default printer: ", "Menus:\n"]
    # iterates through the array, adds c[0] + line to email until len(c)==1 then just adds lines
    for line in user:
        if len(c) > 0:
            message.Body += c.pop(0) + line
        else:
            message.Body += line
    message.Body += "\n#secure#"
    message.Display()


# creates tkinter multiline entry window that takes the user's email and boss's email
def as400_email(email_template):
    window = Tk()
    message = outlook.CreateItemFromTemplate(email_template)
    textbox = Text(window, bg="white", width=35, height=10)
    textbox.insert(tk.END, "User Email First\nBoss's Email If New Employee")
    textbox.pack(pady=8, padx=8)
    textbox.bind("<FocusIn>", temp_text(textbox))
    b = Button(window, text="Submit", command=lambda: [as400write(message, textbox), close(window)])
    b.pack()
    window.mainloop()


# creates a tkinter window to get input for the file location then closes the window
def phone_email(email_template):
    window = Tk()
    message = outlook.CreateItemFromTemplate(email_template)
    textbox = Text(window, bg="white", width=35, height=10)
    textbox.insert(tk.END, "User Email First\nBoss's Email If New Employee")
    textbox.pack(pady=8, padx=8)
    textbox.bind("<FocusIn>", temp_text(textbox))
    b = Button(window, text="Submit", command=lambda: [prop(message, textbox), close(window)])
    b.pack()
    window.mainloop()


# gets email info
def prop(message, textbox):
    user = textbox.get("1.0", 'end').split('\n')
    adressEmail = []
    message.Subject = "Your Company Phone has been ordered."
    for address in user:
        adressEmail.append(address)

    message.To = adressEmail[0]
    if len(adressEmail) > 1:
        message.CC = adressEmail[1]
    message.Display()


# supposed to delete entrance text
def temp_text(textbox):
    textbox.delete("1.0", "end")

#closes window
def close(window):
    window.destroy()


root = tk.Tk()
frame = tk.Frame(root, width=300, height=300)
frame.pack()

outlook = win32.Dispatch('Outlook.Application')
email_templates = ['C:\Template.oft', 'C:\AS400-1.oft', 'C:\PhoneOrder.oft']
label = Label(frame, text="Choose Your Email Template")
label.pack(pady=20)
phone_email_template = tk.Button(frame,
                                 text="AS400",
                                 command=lambda: as400_email(email_templates[1]))
phone_email_template.pack(side=tk.LEFT)
logan = tk.Button(frame,
                  text="Conference",
                  command=lambda: conference_email(email_templates[0]))
logan.pack(side=tk.LEFT)
ogan = tk.Button(frame,
                 text="Phone",
                 command=lambda: phone_email(email_templates[2]))
ogan.pack(side=tk.LEFT)
gan = tk.Button(frame,
                text="Quit",
                fg="red",
                command=quit)
gan.pack(side=tk.LEFT)

root.mainloop()
