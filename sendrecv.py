import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import imaplib
import email
from email.header import decode_header


def send(sending_email, receiver_email, password, subject, content_txt, content_html):
    port = 587  # For starttls
    smtp_server = "smtp-mail.outlook.com"

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sending_email
    message["To"] = receiver_email

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(content_txt, "plain")
    part2 = MIMEText(content_html, "html")

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