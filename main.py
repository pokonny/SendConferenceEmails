import win32com.client as win32
from icalendar import Calendar, Event, vCalAddress, vText
import datetime
from pathlib import Path
import os
import pytz

'''Opens Outlook and Creates an email from template'''
outlook = win32.Dispatch('Outlook.Application')
message = outlook.CreateItemFromTemplate('C:\Template.oft')
# init the calendar
cal = Calendar()

# Some properties are required to be compliant
cal.add('prodid', '-//My calendar product//example.com//')
cal.add('version', '2.0')
filenum = input('num')
while filenum != 'q':
    # Finds the file and reads it
    filenum = input('num')
    msg = r'C:\Users\pokonny\Downloads\calendar('+filenum+').ics'
    e = open(msg, 'rb')
    ecal = cal.from_ical(e.read())

# Walks through the file and adds it to the email
    for component in ecal.walk():
     if component.name == 'VEVENT':
       meetorg = component.get('organizer').replace('mailto:', '')
       meetloc = component.get('location')

       #Checks to see if the file contains a ; for when microsoft teams is included
       if ';' in meetloc:
           meetloca = meetloc.split(';')
           meetloc = meetloca[1]

        #Grabs meeting date and time then formats it in MM/DD/YYYY AND XX:YY AM/PM
       meetdate = component.decoded('dtstart')
       formatted_date = datetime.date.strftime(meetdate, '%m/%d/%Y ')
       formatted_time = datetime.date.strftime(meetdate, '%H:%M %p')

    e.close()
    message.To = meetorg
    spiel = meetloc + ' on ' + formatted_date + 'at ' + formatted_time + '. '
    message.Subject = 'Do you need Tech Assistance at ' + spiel
    message.Body += ' ' + spiel
    message.Display()

