import os
from playsound import playsound
import smtplib
import email
import imaplib
import speech_recognition as sr
from gtts import gTTS
from email.header import decode_header
from typing import Optional
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import imghdr
from email.message import EmailMessage

from CONSTANTS import EMAIL_ID, PASSWORD, LANGUAGE


def SpeakText(command, langinp=LANGUAGE):
    if langinp == "": langinp = "en"
    tts = gTTS(text=command, lang=langinp)
    tts.save("~tempfile01.mp3")
    playsound("~tempfile01.mp3")
    print(command)
    os.remove("~tempfile01.mp3")

def speech_to_text():
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source2:
            r.adjust_for_ambient_noise(source2, duration=0.2)
            audio2 = r.listen(source2)
            MyText = r.recognize_google(audio2)
            print("You said: " + MyText)
            return MyText

    except sr.RequestError as e:
        SpeakText("Could not request results")
        SpeakText(e)
        print("Could not request results; {0}".format(e))
        return None

    except sr.UnknownValueError:
        SpeakText("unknown error occured")
        print("unknown error occured")
        return None


def sendMail(sendTo, msg):
    """
    To send a mail

    Args:
        sendTo (list): List of mail targets
        msg (str): Message
    """
    mail = smtplib.SMTP('smtp.gmail.com', 587)  # host and port
    # Hostname to send for this command defaults to the FQDN of the local host.
    mail.ehlo()
    mail.starttls()  # security connection
    mail.login(EMAIL_ID, PASSWORD)  # login part
    for person in sendTo:
        mail.sendmail(EMAIL_ID, person, msg)  # send part
        print("Mail sent successfully to " + person)
        SpeakText("Mail Sent")
    mail.close()

def composeMail():
    SpeakText("Enter the email IDs of the persons to whom you want to send a mail, separated by 'and': ")
    receivers = speech_to_text()
    receivers = receivers.capitalize()

    email_dict = {
        "Rahul": "rahul.jestadi@gmail.com",
        "Matthew": "leojmatt02@gmail.com",
        "Brinda": "brindadh12@gmail.com",
        "Neha": "neharb8@gmail.com"
        "Shashank": "sha"
    }
    if receivers in email_dict:
        receivers = email_dict[receivers]
        emails = receivers.split(" and ")
        index = 0
        for email in emails:
            emails[index] = email.replace(" ", "")
            index += 1
    else:
        SpeakText(f"Email ID of {receivers} not found. Please mention the email ID.")
        receiver_email = speech_to_text()
        receiver_email = receiver_email.replace("at the rate", "@")
        emails = receivers.split(" and ")
        index = 0
        for email in emails:
            emails[index] = email.replace(" ", "")
            index += 1

    SpeakText("The mail will be send to " +
              (' and '.join([str(elem) for elem in emails])) + ". Confirm by saying YES or NO.")
    confirmMailList = speech_to_text()

    if confirmMailList != "yes":
        SpeakText("Operation cancelled by the user")
        print("Operation cancelled by the user")
        return None

    msg = EmailMessage()
    msg['From'] = EMAIL_ID
    msg['To'] = ', '.join(emails)
    msg['Subject'] = 'Test email'

    SpeakText("Say your message")
    body_text = speech_to_text()
    msg.set_content(body_text)

    SpeakText("Would you like to attach a file? Say YES or NO.")
    attach_file = speech_to_text()
    if attach_file.lower() == "yes":
        SpeakText("Please say the name of the file you want to attach")
        filename = speech_to_text()
        file_path = "/Users/rahuljestadi/Desktop/PROJECTS/WE-MAIL/"+filename
        if not file_path:
            SpeakText(f"{filename} not found in the current directory. Please try again.")
            return None

        with open(file_path, 'rb') as f:
            file_data = f.read()
            file_type = imghdr.what(f.name)
            file_name = f.name
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    SpeakText("You said  " + body_text + ". Confirm by saying YES or NO.")
    confirmMailBody = speech_to_text()
    if confirmMailBody== "yes":
        SpeakText("Message sent")
        msg_str = msg.as_string()
        msg_bytes = msg_str.encode('utf-8')
        sendMail(emails,msg_bytes)
    else:
        SpeakText("Operation cancelled by user")
        print("Operation cancelled by the user")
        return None


def getMailBoxStatus():
    """
    Get mail counts of all folders in the mailbox
    """
    # host and port (ssl security)
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    M.login(EMAIL_ID, PASSWORD)  # login

    for i in M.list()[1]:
        l = i.decode().split(' "/" ')
        if l[1] == '"[Gmail]"':
            continue

        stat, total = M.select(f'{l[1]}')
        l[1] = l[1][1:-1]
        messages = int(total[0])
        if l[1] == 'INBOX':
            SpeakText(l[1] + " has " + str(messages) + " messages.")
        else:
            SpeakText(l[1].split("/")[-1] + " has " + str(messages) + " messages.")

    M.close()
    M.logout()


def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)


def getLatestMails():
    mailBoxTarget = "INBOX"
    SpeakText(
        "Choose the folder name to get the latest mails. Say inbox for Inbox. Say sent mail for Sent Mailbox. Say "
        "drafts for Drafts. "
        "Say important for important mails. Say spam for Spam. Say starred for Starred Mails. Say bin for Bin.")
    cmb = speech_to_text()
    if cmb == "1" or cmb.lower() == "one" or cmb.lower() == "inbox" or cmb.lower() == "inbox mail":
        mailBoxTarget = "INBOX"
        SpeakText("Inbox selected.")
    elif cmb == "2" or cmb.lower() == "two" or cmb.lower() == "tu" or cmb.lower() == "sent mail" or cmb.lower() == "sent" or cmb.lower() == "one":
        mailBoxTarget = '"[Gmail]/Sent Mail"'
        SpeakText("Sent Mailbox selected.")
    elif cmb == "3" or cmb.lower() == "three" or cmb.lower() == "drafts" or cmb.lower() == "draft mail" or cmb.lower() == "draft":
        mailBoxTarget = '"[Gmail]/Drafts"'
        SpeakText("Drafts selected.")
    elif cmb == "4" or cmb.lower() == "four" or cmb.lower() == "important" or cmb.lower() == "important mail" or cmb.lower() == "important mailbox":
        mailBoxTarget = '"[Gmail]/Important"'
        SpeakText("Important Mails selected.")
    elif cmb == "5" or cmb.lower() == "five" or cmb.lower() == "spam mail" or cmb.lower() == "spam mailbox" or cmb.lower() == "spam":
        mailBoxTarget = '"[Gmail]/Spam"'
        SpeakText("Spam selected.")
    elif cmb == "6" or cmb.lower() == "six" or cmb.lower() == "starred" or cmb.lower() == "starred mail":
        mailBoxTarget = '"[Gmail]/Starred"'
        SpeakText("Starred Mails selected.")
    elif cmb == "7" or cmb.lower() == "seven" or cmb.lower() == "bin" or cmb.lower() == "mail bin":
        mailBoxTarget = '"[Gmail]/Bin"'
        SpeakText("Bin selected.")
    else:
        SpeakText("Wrong choice. Hence, default option Inbox wil be selected.")

    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(EMAIL_ID, PASSWORD)

    status, messages = imap.select(mailBoxTarget)

    messages = int(messages[0])

    if messages == 0:
        SpeakText("Selected MailBox is empty.")
        return None
    elif messages == 1:
        N = 1  # number of top emails to fetch
    elif messages == 2:
        N = 2  # number of top emails to fetch
    else:
        N = 3  # number of top emails to fetch

    msgCount = 1
    for i in range(messages, messages - N, -1):
        SpeakText(f"Message {msgCount}:")
        res, msg = imap.fetch(str(i), "(RFC822)")  # fetch the email message by ID
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])  # parse a bytes email into a message object

                subject, encoding = decode_header(msg["Subject"])[0]  # decode the email subject
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)  # if it's a bytes, decode to str

                From, encoding = decode_header(msg.get("From"))[0]  # decode email sender
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                SpeakText("Subject: " + subject)
                FromArr = From.split()
                FromName = " ".join(namechar for namechar in FromArr[0:-1])
                SpeakText("From: " + FromName)
                SpeakText("Sender mail: " + FromArr[-1])
                SpeakText("The mail says or contains the following:")

                # MULTIPART
                if msg.is_multipart():
                    for part in msg.walk():  # iterate over email parts
                        content_type = part.get_content_type()  # extract content type of email
                        content_disposition = str(
                            part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()  # get the email body
                        except:
                            pass

                        # PLAIN TEXT MAIL
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            SpeakText("Do you want to listen to the text content of the mail ? Please say YES or NO.")
                            talkMSG1 = speech_to_text()
                            if "yes" in talkMSG1.lower():
                                SpeakText("The mail body contains the following:")
                                SpeakText(body)
                            else:
                                SpeakText("You chose NO")

                        # MAIL WITH ATTACHMENT
                        elif "attachment" in content_disposition:
                            SpeakText(
                                "The mail contains attachment, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail")
                            filename = part.get_filename()  # download attachment
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)  # make a folder for this email (named after the subject)
                                filepath = os.path.join(folder_name, filename)
                                open(filepath, "wb").write(
                                    part.get_payload(decode=True))  # download attachment and save it

                # NOT MULTIPART
                else:
                    content_type = msg.get_content_type()  # extract content type of email
                    body = msg.get_payload(decode=True).decode()  # get the email body
                    if content_type == "text/plain":
                        SpeakText("Do you want to listen to the text content of the mail ? Please say YES or NO.")
                        talkMSG2 = speech_to_text()
                        if "yes" in talkMSG2.lower():
                            SpeakText("The mail body contains the following:")
                            SpeakText(body)
                        else:
                            SpeakText("You chose NO")

                # HTML CONTENTS
                if content_type == "text/html":
                    SpeakText(
                        "The mail contains an HTML part, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail. You can view the html files in any browser, simply by clicking on them.")
                    # if it's HTML, create a new HTML file
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        os.mkdir(folder_name)  # make a folder for this email (named after the subject)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)

                    # webbrowser.open(filepath)     # open in the default browser

                SpeakText(f"\nEnd of message {msgCount}:")
                msgCount += 1
                print("=" * 100)
    imap.close()
    imap.logout()


def searchMail():
    """
    Search mails by subject / author mail ID

    Returns:
        None: None
    """
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    M.login(EMAIL_ID, PASSWORD)

    mailBoxTarget = "INBOX"
    SpeakText(
        "Where do you want to search ? Say inbox for Inbox. Say sent mailbox for Sent Mailbox. Say drafts for Drafts. "
        "Say important mails for important mails. Say five for Spam. Say six for Starred Mails. Say bin for Bin.")
    cmb = speech_to_text()
    if cmb == "1" or cmb.lower() == "one" or cmb.lower() == "inbox" or cmb.lower() == "inbox mail":
        mailBoxTarget = "INBOX"
        SpeakText("Inbox selected.")
    elif cmb == "2" or cmb.lower() == "two" or cmb.lower() == "tu" or cmb.lower() == "sent" or cmb.lower() == "sent mail":
        mailBoxTarget = '"[Gmail]/Sent Mail"'
        SpeakText("Sent Mailbox selected.")
    elif cmb == "3" or cmb.lower() == "three" or cmb.lower() == "drafts" or cmb.lower() == "draft" or cmb.lower() == "draft mail":
        mailBoxTarget = '"[Gmail]/Drafts"'
        SpeakText("Drafts selected.")
    elif cmb == "4" or cmb.lower() == "four" or cmb.lower() == "important" or cmb.lower() == "important mail":
        mailBoxTarget = '"[Gmail]/Important"'
        SpeakText("Important Mails selected.")
    elif cmb == "5" or cmb.lower() == "five" or cmb.lower() == "spam" or cmb.lower() == "spam folder" or cmb.lower() == "spam mail":
        mailBoxTarget = '"[Gmail]/Spam"'
        SpeakText("Spam selected.")
    elif cmb == "6" or cmb.lower() == "six" or cmb.lower() == "starred mail" or cmb.lower() == "starred folder" or cmb.lower() == "starred mail":
        mailBoxTarget = '"[Gmail]/Starred"'
        SpeakText("Starred Mails selected.")
    elif cmb == "7" or cmb.lower() == "seven" or cmb.lower() == "bin" or cmb.lower() == "bin folder":
        mailBoxTarget = '"[Gmail]/Bin"'
        SpeakText("Bin selected.")
    else:
        SpeakText("Wrong choice. Hence, default option Inbox wil be selected.")

    M.select(mailBoxTarget)

    SpeakText(
        "Say sender to search mails from a specific sender. Say subject to search mail with respect to the subject of the mail.")
    mailSearchChoice = speech_to_text()
    if mailSearchChoice == "sender" or mailSearchChoice.lower() == "sender mail" or mailSearchChoice.lower() == "email" or mailSearchChoice.lower() == "by sender" or mailSearchChoice.lower() == "sendor":
        SpeakText("Please mention the sender email ID you want to search.")
        searchSub = speech_to_text()
        searchSub = searchSub.replace("at the rate", "@")
        searchSub = searchSub.replace(" ", "")
        status, messages = M.search(None, f'FROM "{searchSub}"')
    elif mailSearchChoice == "2" or mailSearchChoice.lower() == "two" or mailSearchChoice.lower() == "tu" or mailSearchChoice.lower() == "subject" or mailSearchChoice.lower() == "mail subject":
        SpeakText("Please mention the subject of the mail you want to search.")
        searchSub = speech_to_text()
        status, messages = M.search(None, f'SUBJECT "{searchSub}"')
    else:
        SpeakText(
            "Wrong choice. Performing default operation. Please mention the subject of the mail you want to search.")
        searchSub = speech_to_text()
        status, messages = M.search(None, f'SUBJECT "{searchSub}"')

    if str(messages[0]) == "b''":
        SpeakText(f"Mail not found in {mailBoxTarget}.")
        return None

    msgCount = 1
    for i in messages:
        SpeakText(f"Message {msgCount}:")
        res, msg = M.fetch(i, "(RFC822)")  # fetch the email message by ID
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])  # parse a bytes email into a message object

                subject, encoding = decode_header(msg["Subject"])[0]  # decode the email subject
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)  # if it's a bytes, decode to str

                From, encoding = decode_header(msg.get("From"))[0]  # decode email sender
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                SpeakText("Subject: " + subject)
                FromArr = From.split()
                FromName = " ".join(namechar for namechar in FromArr[0:-1])
                SpeakText("From: " + FromName)
                SpeakText("Sender mail: " + FromArr[-1])

                # MULTIPART
                if msg.is_multipart():
                    for part in msg.walk():  # iterate over email parts
                        content_type = part.get_content_type()  # extract content type of email
                        content_disposition = str(
                            part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()  # get the email body
                        except:
                            pass

                        # PLAIN TEXT MAIL
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            SpeakText("Do you want to listen to the text content of the mail ? Please say YES or NO.")
                            talkMSG1 = speech_to_text()
                            if "yes" in talkMSG1.lower():
                                SpeakText("The mail body contains the following:")
                                SpeakText(body)
                            else:
                                SpeakText("You chose NO")

                        # MAIL WITH ATTACHMENT
                        elif "attachment" in content_disposition:
                            SpeakText(
                                "The mail contains attachment, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail")
                            filename = part.get_filename()  # download attachment
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)  # make a folder for this email (named after the subject)
                                filepath = os.path.join(folder_name, filename)
                                open(filepath, "wb").write(
                                    part.get_payload(decode=True))  # download attachment and save it

                # NOT MULTIPART
                else:
                    content_type = msg.get_content_type()  # extract content type of email
                    body = msg.get_payload(decode=True).decode()  # get the email body
                    if content_type == "text/plain":
                        SpeakText("Do you want to listen to the text content of the mail ? Please say YES or NO.")
                        talkMSG2 = speech_to_text()
                        if "yes" in talkMSG2.lower():
                            SpeakText("The mail body contains the following:")
                            SpeakText(body)
                        else:
                            SpeakText("You chose NO")

                # HTML CONTENTS
                if content_type == "text/html":
                    SpeakText(
                        "The mail contains an HTML part, the contents of which will be saved in respective folders "
                        "with it's name similar to that of subject of the mail. You can view the html files in any "
                        "browser, simply by clicking on them.")
                    # if it's HTML, create a new HTML file
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        os.mkdir(folder_name)  # make a folder for this email (named after the subject)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)

                    # webbrowser.open(filepath)     # open in the default browser

                SpeakText(f"\nEnd of message {msgCount}:")
                msgCount += 1
                print("=" * 100)
    M.close()
    M.logout()


def get_latest_mail_by_sender(sender: str) -> Optional[Message]:
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    M.login(EMAIL_ID, PASSWORD)

    M.select("INBOX")
    status, data = M.search(None, "FROM", f'"{sender}"')
    mail_ids = data[0].split()

    if not mail_ids:
        SpeakText(f"No email found from {sender}")
        return None

    latest_mail_id = mail_ids[-1]  # get the latest email
    status, data = M.fetch(latest_mail_id, "(RFC822)")
    mail_data = data[0][1]
    mail_message = email.message_from_bytes(mail_data)

    return mail_message


def get_latest_mail_by_subject(subject: str) -> Optional[Message]:
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    M.login(EMAIL_ID, PASSWORD)

    M.select("INBOX")
    status, data = M.search(None, "SUBJECT", f'"{subject}"')
    mail_ids = data[0].split()

    if not mail_ids:
        SpeakText(f"No email found with subject {subject}")
        return None

    latest_mail_id = mail_ids[-1]  # get the latest email
    status, data = M.fetch(latest_mail_id, "(RFC822)")
    mail_data = data[0][1]
    mail_message = email.message_from_bytes(mail_data)

    return mail_message


def reply_to_mail(sender=None, subject=None):
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    M.login(EMAIL_ID, PASSWORD)

    M.select("INBOX")
    search_criteria = ['ALL']
    if sender:
        search_criteria.append(f'FROM "{sender}"')
    if subject:
        search_criteria.append(f'SUBJECT "{subject}"')
    search_criteria = " ".join(search_criteria)
    status, data = M.search(None, search_criteria)
    mail_ids = data[0].split()

    if not mail_ids:
        SpeakText("No emails found with the specified search criteria.")
        return None

    latest_mail_id = mail_ids[-1]  # get the latest email
    status, data = M.fetch(latest_mail_id, "(RFC822)")
    mail_data = data[0][1]
    mail_message = email.message_from_bytes(mail_data)

    From, encoding = decode_header(mail_message.get("From"))[0]
    if isinstance(From, bytes):
        From = From.decode(encoding)

    subject, encoding = decode_header(mail_message["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    SpeakText(f"Replying to email from {From} with subject {subject}.")

    reply_message = MIMEMultipart()
    reply_message["In-Reply-To"] = mail_message["Message-ID"]
    reply_message["References"] = mail_message["Message-ID"]
    reply_message["To"] = mail_message["Reply-To"] or mail_message["From"]
    reply_message["Subject"] = f"Re: {subject}"
    reply_message["From"] = EMAIL_ID

    body_text = "Say the reply message you want to send."
    SpeakText(body_text)
    reply_text = speech_to_text()

    text_part = MIMEText(reply_text, "plain")
    reply_message.attach(text_part)

    try:
        M.append("INBOX", None, imaplib.Time2Internaldate(time.time()), reply_message.as_bytes())
        SpeakText("Reply sent successfully.")
    except:
        SpeakText("Error while sending reply.")


def main():
    if EMAIL_ID != "" and PASSWORD != "":

            SpeakText("Choose and speak out the option number for the task you want to perform. Say compose to send a "
                    "mail. Say status to get your mailbox status. Say search to search a mail. Say last three to get the "
                    "last 3 mails. Say reply to reply to a mail")
            choice = speech_to_text()
            if choice == 'compose mail' or choice == 'compose' or choice == 'make' or choice == 'send' or choice == 'send mail':
                composeMail()
            elif choice == 'status' or choice == 'mailbox status' or choice == "stat us":
                getMailBoxStatus()
            elif choice == 'search' or choice == 'search mail' or choice == 'mail search':
                searchMail()
            elif choice == 'recent' or choice == 'last' or choice == 'last three' or choice == "sent" or choice == "last 3":
                getLatestMails()
            elif choice == 'reply' or choice == 'reply to':
                reply_to_mail()
            elif choice == 'stop' or choice == 'exit':
                SpeakText("Not doing anything")

            else:
                SpeakText("Wrong choice. Please say only the appropriate words")
            
            

    else:
        SpeakText("Both Email ID and Password should be present")


if __name__ == '__main__':
    main()
