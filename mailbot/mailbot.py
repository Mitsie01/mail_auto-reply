import sendrecv

import time
import datetime
import sys
import os
from dotenv import load_dotenv

#Create a .env file with your e-mail credentials!

load_dotenv()

SENDING_ADDRESS = os.getenv('USER')
PASSWORD = os.getenv('PASSWORD')


##insert SENDING_ADDRESS, subject and mail content

#host e-mail information
folder = "INBOX"


#content for reply e-mail
send_sub = 'My subject'

send_txt = """\
My email text"""
send_html = """\
<html>
<body>
    <br>
    <h3>My email text</h3><br>
</body>
</html>
"""

#variable
n = 0

#store previous e-mail subject for data retention
try:
    with open('mailbot/last-email.txt') as f:
        previous = f.readline()
    f.close()
except:
    previous = ' '

#scan for e-mail every second
while True:

    #check last received e-mail
    subject, From, value, content = sendrecv.receive(SENDING_ADDRESS, PASSWORD, folder)

    #check if e-mail is new
    if subject == previous:
        read = True
        if n == 4:
            n = 0
        print('Checking'+'.'*n, end='\r')
        sys.stdout.write("\033[K")
        n = n + 1
    else:
        read = False
        n = 0

    #send reply if e-mail is new
    if value and not read:

        #retreive address from sender and check if e-mail is forwarded
        address, forwarded = sendrecv.isolate(From, content)

        #print sender address and e-mail information
        now = datetime.datetime.now()
        date_time = now.strftime("%d/%m/%Y, %H:%M:%S")
        print(f"[{forwarded}][{date_time}] {address}")

        #send reply with prevention from sending to own address
        if address == SENDING_ADDRESS:
            print('[RETURN] Own e-mail returned. Not sending a reply.')
        else:
            sendrecv.send(SENDING_ADDRESS, PASSWORD, address, send_sub, send_txt, send_html) 

        #update last e-mail subject and store in .txt file    
        previous = subject
        with open('mailbot/last-email.txt', 'w') as f:
            f.write(previous)

    time.sleep(1)
