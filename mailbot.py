import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import imaplib
import email
from email.header import decode_header
import time
import datetime

import getpass
import sys


#insert SENDING_ADDRESS, subject and mail content



def send(receiver_email, password, sending_email):
    port = 587  # For starttls
    smtp_server = "smtp-mail.outlook.com"
    subject = "subject"

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sending_email
    message["To"] = receiver_email

    # Create the plain-text and HTML version of your message
    text = """\
    My email text"""
    html = """\
    <html>
    <body>
        <br>
        <h3>My email text</h3><br>
    </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.login(sending_email, password)
        server.sendmail(
            sending_email, receiver_email, message.as_string()
        )


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def receive(password, folder, sending_address):
    value = False
    subject = 'Foo'
    From = 'Bar'
    body = 'cheese'

    imap_server = "outlook.office365.com" #port 993

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL(imap_server)
    # authenticate
    imap.login(sending_address, password)

    status, messages = imap.select(folder)
    # number of top emails to fetch
    N = 1
    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages-N, -1):
    # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                # print("Subject:", subject)
                # print("From:", From)
                value = True
                

                if msg.is_multipart():
                # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            content = body
                            # print(body)

                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    payload = msg.get_payload(decode=True)
                    if payload is None:
                        body = 'cheese'
                        continue
                    body = payload.decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        content = body
                        # print(body)
    imap.close()
    imap.logout()

    return subject, From, value, content

def isolate(address, content):
    try:
        check = content.split('---------')
        if "- Forwarded message " in check:
            addresses = content.split('<')
            address = addresses[1].split()
            address = address[0].replace('>', '')
            forwarded = 'FORWARDED'
        else:
            elements = address.split('<')
            address = elements[len(elements) - 1]
            address = address.replace('>', '')
            forwarded = 'DIRECT'
            
    except:
        elements = address.split('<')
        address = elements[len(elements) - 1]
        address = address.replace('>', '')
        forwarded = 'DIRECT'

    return address, forwarded




try:
    with open('last-email.txt') as f:
        previous = f.readline()
    f.close()
except:
    previous = ' '

SENDING_ADDRESS = 'my.email@address.com'
password = getpass.getpass("Password: ")
folder = "INBOX"
n = 0

while True:

    subject, From, value, content = receive(password, folder, SENDING_ADDRESS)

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

    if value and not read:
        address, forwarded = isolate(From, content)
        now = datetime.datetime.now()
        date_time = now.strftime("%d/%m/%Y, %H:%M:%S")
        print(f"[{forwarded}][{date_time}] {address}")
        if address == SENDING_ADDRESS:
            print('[RETURN] Own e-mail returned. Not sending a reply.')
        else:
            send(address, password, SENDING_ADDRESS) 
        previous = subject
        with open('last-email.txt', 'w') as f:
            f.write(previous)

    time.sleep(1)
