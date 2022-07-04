from __future__ import print_function
from email.message import EmailMessage
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import date, datetime, timedelta
import os.path
import base64

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
# Fetches service.users().messages() as infrequently as possible by initializing
# as global variable
msgs = None

def main():
    """ Collects emails sent from EMAILS and sends them once every three days in
    a digest to prevent inbox clutter.
    """
    global msgs 
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        msgs = service.users().messages()

        # Only retrieve emails from the label Unwanted/For Digest
        # and from 7 days ago to now.
        currDate = date.today()
        tmrw = (currDate + timedelta(days = 1)).strftime("%y/%m/%d")
        daysBefore = (currDate - timedelta(days=7)).strftime("%y/%m/%d")
        query = f"label:unwanted-for-digest before:{tmrw} after:{daysBefore}"
        results = msgs.list(userId='me', q = query).execute()
        
        # if there aren't any emails to be compiled, stop
        if results['resultSizeEstimate'] == 0:
            print("There are no emails")
            return
        digestMsg = createEmails(results['messages'])
        sendEmail(digestMsg)
    except HttpError as error:
        print(f'An error occurred: {error}')

def createEmails(emails):
    """ Takes a list of emails and extracts the contents of each email to
    combine into one, big digest
    """

    # a table of contents that goes in the beginning of the digest
    digestHeader = """
    <h1><a name='toc'>contents of digest<a/></h1>
    <ol>

    """
    digest = ""
    # keeps track of indices for table of contents
    count = 0

    for mail in emails:
        # extract body of email
        content = msgs.get(userId = 'me', id = mail['id'], format =
    "full").execute()['payload']
        data = ""
        # if multipart, then extract html text
        if content["mimeType"] == "multipart/alternative":
            for part in content["parts"]:
                if part["mimeType"] == "text/html":
                    data = part["body"]["data"]
                    break
        # otherwise, it doesn't have any parts so
        # can read data directly
        else:
            data  = content["body"]["data"]
        headers = content["headers"]
        subject, sender, dateSent  = "", "", ""
        for h in headers:
            name = h["name"]
            val = h["value"]
            if name == "Subject":
                subject = val
            elif name == "Reply-To":
                sender = val#[:val.index("<")]
            elif name == "Date":
                dateSent = val
        digestHeader += f"<li><a href='#{count}'>{sender}: {subject}</a></li>\n"
            
            # add header that will be used for navigation via table of contents
        digest += f"""
            <h2 style = 'display:inline'><a name={count}>{sender}: {subject}</a></h2>
            <p style='color:gray'>from: {dateSent}</p>
            <h3><a href='#toc'>return to toc</a></h3>
        """
        # add contents of email to digest
        digest += base64.urlsafe_b64decode(
            data.encode("ASCII")).decode("utf-8")

            # move email to trash
            #msgs.trash(userId = 'me', id = mail['id']).execute()
        count += 1
            
    digestHeader += """
    </ol>
    """

    digest = digestHeader + digest
    
    # Create email
    message = EmailMessage()
    message.set_content(digest)
    message["To"] = "lglee@g.hmc.edu"
    message["From"] = "leelillian205@gmail.com"

    # Formatting subject of email
    currDate = date.today()
    dayBefore = (currDate + timedelta(days = 1)).strftime("%m/%d/%y")
    weekBefore = (currDate - timedelta(days=60)).strftime("%m/%d/%y")
    message["Subject"] = f"Weekly Digest {dayBefore}~{weekBefore}"

    message.add_alternative(digest, subtype="html")
    return message

def sendEmail(message):
    """ Takes EmailMessage and sends it to the receiver """
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    create_message = {'raw': encoded_message}
    send_message = msgs.send(userId = "me", body = create_message).execute()
    print("Email has been sent!")


if __name__ == '__main__':
    main()
