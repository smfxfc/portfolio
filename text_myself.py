#! python3
# textmyself.py - Defines the text_myself) function that texts a message
# passed to it as a string

# Preset values:
accountSID = '#########'
authToken = '#########'
myNumber = '#########'
twilioNumber = '#########'

from twilio.rest import Client

def textmyself(message):
    twilioCli = Client(accountSID, authToken)
    twilioCli.messages.create(body=message, from_=twilioNumber, to=myNumber)

