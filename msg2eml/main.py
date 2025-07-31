import extract_msg
from email.message import EmailMessage
from email.policy import SMTP
import os
from rich.console import Console
import sys
import argparse
import mimetypes


def convert_msg_to_eml(msg_path, eml_path=None):
    # use the same file name and location by default
    if eml_path is None:
        base, _ = os.path.splitext(msg_path)
        eml_path = base + ".eml"

    # pull data from email object (.msg)
    msg = extract_msg.Message(msg_path)
    msg_sender = msg.sender
    msg_recipients = msg.to
    msg_cc = msg.cc
    msg_bcc = msg.bcc
    msg_subject = msg.subject
    msg_date = msg.date
    msg_body = msg.body
    msg_html = msg.htmlBody
    
    # create new email object (.eml)
    email = EmailMessage(policy=SMTP)

    # check field exists before assignment
    if msg_sender: email['From'] = msg_sender
    if msg_recipients: email['To'] = msg_recipients
    if msg_cc: email['Cc'] = msg_cc
    if msg_bcc: email['Bcc'] = msg_bcc
    if msg_subject: email['Subject'] = msg_subject
    if msg_date: email['Date'] = msg_date
    
    # write content to email object (.eml)    
    if msg_html:
        if isinstance(msg_html, bytes):
            msg_html = msg_html.decode(errors='replace')

        email.set_content(msg_body or " ")
        email.add_alternative(msg_html, subtype='html')
    else:
        email.set_content(msg_body or " ")
    
    # attach files
    for attachment in msg.attachments:
        filename = attachment.longFilename or attachment.shortFilename or "attachment"
        mime_type, _ = mimetypes.guess_type(filename)
        
        if mime_type:
            maintype, subtype = mime_type.split("/")
        else:
            # fallback if no mime type identified
            maintype, subtype = "application", "octet-stream"
        
        email.add_attachment(attachment.data, maintype=maintype, subtype=subtype, filename=filename)

    # write
    with open(eml_path, 'wb') as f:
        f.write(email.as_bytes())    

