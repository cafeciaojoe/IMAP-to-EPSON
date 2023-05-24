import imaplib
import email
from email.header import decode_header
from email.utils import parsedate
import webbrowser
import os

import datetime
import time

# account credentials
username = "cheapestpartydjs@hotmail.com"
password = "Pic[17]asso"
# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
# for office 365, it's this:
imap_server = "outlook.office365.com"

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def print_check(date):
    """ datetime.datetime A combination of a date and a time.
    Attributes: year, month, day, hour, minute, second, microsecond, and tzinfo."""

    """https://docs.python.org/3/library/email.utils.html"""

    """time zone not passed in here, did not want to chang hot mail setting in case they took a long time to propogate
    so i just added +2 to the hour argument"""

    with open("latest_unix.txt", "r") as my_file:
        read_unix = int(my_file.read())
        print("read_unix_value =", read_unix)
        my_file.close()

    date_parsed = email.utils.parsedate(date)
    date_time = datetime.datetime(
        date_parsed[0],
        date_parsed[1],
        date_parsed[2],
        date_parsed[3]+2,
        date_parsed[4],
        date_parsed[6])
    passed_in_unix = int((time.mktime(date_time.timetuple())))

    print("Given Date:", date_time)
    print("UNIX timestamp:",passed_in_unix)

    if read_unix >= passed_in_unix:
        print("Print check ture, email has already been printed")
        return True
    else:
        print("Print check false, email has not been printed yet")
        with open("latest_unix.txt", "w") as my_file:
                my_file.write(str(passed_in_unix))
                print("writing passed_in_unix to file")
                my_file.close()
        return False

#print_check("Fri, 19 May 2023 19:46:27 -0700")
#print("hi")

def thermal_print(From,subject, body):
    max_body_lengh = 1000 #characters

    print("BZZZZZZZZ"*100)
    print("FROM", From)
    print("SUBJECT",subject)
    print("BODY", body[:max_body_lengh])
    print("BZZZZZZZZ"*100)


# create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
imap.login(username, password)

status, messages = imap.select("INBOX")
# number of top emails to fetch
#N = 3
# total number of emails
messages = int(messages[0])
print(messages)
"""added plus one to for loop range so it does not go to 0,that causes an error The specified message set is invalid"""
for i in range(1,messages+1, +1):
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
            print("Subject:", subject)
            print("From:", From)
            # if the email message is multipart
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
                        print(body)
                        thermal_body = body
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            folder_name = clean(subject)
                            if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                os.mkdir(folder_name)
                            filepath = os.path.join(folder_name, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
                    thermal_body = body
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                """
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                webbrowser.open(filepath)
                """
            print("="*100)
    # print(msg.keys())
    if print_check(msg["date"]) is False:
        thermal_print(From,subject,thermal_body)

# close the connection and logout
imap.close()
imap.logout()
