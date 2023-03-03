import win32com.client as win32
import tkinter as tk
from tkinter import *


class EmailStart:

    def __init__(self):
        self.window = None
        self.button = None
        self.message = None

    def __int__(self):
        self.message = None

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

    # takes the input then puts it into an array, opens outlook, assigns first two values in email_to_cc to the To/CC
    # field then removes them
    def email_create(self, text_box, choice_string):
        email_to_cc = text_box.get("1.0", 'end').split('\n')
        outlook = win32.Dispatch('Outlook.Application')
        self.message = outlook.CreateItemFromTemplate(choice_string)
        self.message.To = email_to_cc.pop(0)
        self.message.CC = email_to_cc.pop(0)
        if len(email_to_cc) <= 1:
            self.message.Subject = "Your Company Phone Has Been Ordered."
            self.message.Display()
        else:
            self.message.Subject = "Your AS400 access has been granted."
            self.as400_write(email_to_cc)

    # creates a multi-line entry window then, on button click, calls the email_create class then destroys itself
    def info_enter(self, choice_string):
        self.window = Tk()
        textbox = Text(self.window, bg="white", width=35, height=10)
        textbox.pack(pady=8, padx=8)
        b = Button(self.window, text="Submit", command=lambda: [self.email_create(textbox, choice_string),
                                                                self.window.destroy()])
        b.pack()
        self.window.mainloop()

    # creates the selector buttons that calls info_enter
    def button_choice(self, frame_grab, tk_grab, choice_string, action_to_take):
        if action_to_take == "quit":
            self.button = tk.Button(frame_grab,
                                    fg='red',
                                    text=choice_string,
                                    command=quit)
            self.button.pack(side=tk_grab.LEFT, pady=10, padx=10)
        else:
            self.button = tk.Button(frame_grab,
                                    text=choice_string,
                                    command=lambda: self.info_enter(action_to_take))
            self.button.pack(side=tk_grab.LEFT, pady=10, padx=10)


create_email_from_template = EmailStart()
root = tk.Tk()
frame = tk.Frame(root, width=300, height=300)
frame.pack()
label = Label(frame, text="Choose Your Email Template")
label.pack(pady=20)
create_email_from_template.button_choice(frame, tk, "AS400", r'C:\AS400-1.oft')
create_email_from_template.button_choice(frame, tk, "Phone", r'C:\PhoneOrder.oft')
create_email_from_template.button_choice(frame, tk, "Quit", "quit")
label.pack(pady=20)
root.mainloop()
