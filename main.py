import tkinter as tk
from tkinter import *
import win32com.client as win32
from icalendar import Calendar, Event, vCalAddress, vText
import datetime


class EmailStart:

    def __init__(self):
        self.window = None
        self.button = None
        self.message = None
        self.email_template = r'C:\Template.oft'
        self.outlook = win32.Dispatch('Outlook.Application')

    def __int__(self):
        self.message = None

    # creates the selector buttons that calls info_enter
    def button_choice(self, frame_grab, tk_grab, choice_string, action_to_take):
        self.button = tk.Button(frame_grab,
                                text=choice_string,
                                command=lambda: self.info_enter(action_to_take))
        self.button.pack(side=tk_grab.LEFT, pady=10, padx=10)

    # creates a multi-line entry window then, on button click, calls the email_create class then destroys itself
    def info_enter(self, choice_string):
        if choice_string == "quit":
            return quit()
        elif choice_string == "C:\Template.oft":
            self.conference_response()
        else:
            self.window = Tk()
            textbox = Text(self.window, bg="white", width=35, height=10)
            textbox.pack(pady=8, padx=8)
            b = Button(self.window, text="Submit", command=lambda: [self.email_create(textbox, choice_string),
                                                                    self.window.destroy()])
            b.pack()
            self.window.mainloop()

    # takes the input then puts it into an array, opens outlook, assigns first two values in email_to_cc to the To/CC
    # field then removes them
    def email_create(self, text_box, choice_string):
        email_to_cc = text_box.get("1.0", 'end').split('\n')
        self.message = self.outlook.CreateItemFromTemplate(choice_string)
        self.message.To = email_to_cc.pop(0)
        self.message.CC = email_to_cc.pop(0)
        if len(email_to_cc) <= 1:
            self.message.Subject = "Your Company Phone Has Been Ordered."
            self.message.Display()
        else:
            self.message.Subject = "Your AS400 access has been granted."
            self.as400_write(email_to_cc)

    # fills out AS400 login details then displays email
    def as400_write(self, as400_info):
        self.message.Body += "\n"
        c = ["Username: ", "Password: ", "AS400 Default printer: ", "Menus:\n"]
        # iterates through the array, adds c[0] + line to email until len(c)==1 then just adds lines
        for line in as400_info:
            if len(c) > 0:
                self.message.Body += c.pop(0) + line
            else:
                self.message.Body += line
        self.message.Body += "\n#secure#"
        self.message.Display()

    def conference_response(self):
        self.window = Tk()
        label = Label(self.window, text="File location")
        label.pack(side=LEFT)
        file_name = Entry(self.window, bd=5)
        file_name.pack(side=LEFT)
        b = Button(self.window, text="Submit",
                   command=lambda: [self.read_invite(file_name.get(), self.email_template), self.window.destroy()])
        b.pack()
        self.window.mainloop()

    def read_invite(self, file_name, email_template):
        # asks for ics file location opens ics file
        cal = Calendar()
        conf_email_temp = self.outlook.CreateItemFromTemplate(email_template)
        # Finds the file and reads what kind it is
        msg = r'C:\Users\pokonny\Downloads\calendar(' + file_name + ').ics'
        e = open(msg, 'rb')
        ecal = cal.from_ical(e.read())

        # Walks through the file and adds it to the email
        for component in ecal.walk():
            if component.name == 'VEVENT':
                conf_email_temp.To = component.get('organizer').replace('mailto:', '')
                meet_location = component.get('location')
                meet_date = component.decoded('dtstart')
                meet_location = meet_location.replace("Microsoft Teams Meeting;","")

                # Grabs meeting date and time then formats it in MM/DD/YYYY AND XX:YY AM/PM
                formatted_date = datetime.date.strftime(meet_date, '%m/%d/%Y ')
                formatted_time = datetime.date.strftime(meet_date, '%H:%M %p')
                spiel = meet_location + ' on ' + formatted_date + 'at ' + formatted_time + '. '
                conf_email_temp.Subject = 'Do you need Tech Assistance at ' + spiel
                conf_email_temp.Body += ' ' + spiel

        e.close()
        conf_email_temp.Display()


create_email_from_template = EmailStart()
root = tk.Tk()
frame = tk.Frame(root, width=300, height=300)
frame.pack()
label = Label(frame, text="Choose Your Email Template")
label.pack(pady=20)
create_email_from_template.button_choice(frame, tk, "AS400", r'C:\AS400-1.oft')
create_email_from_template.button_choice(frame, tk, "Phone", r'C:\PhoneOrder.oft')
create_email_from_template.button_choice(frame, tk, "Conference", r'C:\Template.oft')
create_email_from_template.button_choice(frame, tk, "Quit", "quit")
label.pack(pady=20)
root.mainloop()
