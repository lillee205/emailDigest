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
EMAILS = ["hello@smart-shopping.thredup.com", "news@omocat.com",
          "messages+5jv970x069v5s@squaremktg.com", "rtuynman@ryman.org"]
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

        # Only retrieve emails from senders in EMAILS and from certain time range
        currDate = date.today()
        tmrw = (currDate + timedelta(days = 1)).strftime("%y/%m/%d")
        daysBefore = (currDate - timedelta(days=60)).strftime("%y/%m/%d")
        query = " ".join(["from:" + _ + " OR " for _ in EMAILS])[:-3] + f"before:{tmrw} after:{daysBefore}"
        print(query)
        results = msgs.list(userId='me', q = query).execute()
        
        # if there aren't any emails to be compiled, stop
        if results['resultSizeEstimate'] == 0:
            print("no messages")
            return
        createEmails(results['messages'])

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
        if 'data' in content["body"]:
            print(content)
            subject = [h["value"] for h in content["headers"] if h["name"] ==
    "Subject"][0]
            sender = [h["value"] for h in content["headers"] if h["name"] ==
    "Reply-To"][0]
            sender = sender[:sender.index("<")]
            digestHeader += f"<li><a href='#{count}'>{sender}: {subject}</a></li>"
            
            # add header that will be used for navigation via table of contents
            digest += f"<h2><a name={count}>{sender}: {subject}</a></h2>\n"
            digest += f"<h3><a href='#toc'>return to toc</a></h3>"
            # add contents of email to digest
            digest += base64.urlsafe_b64decode(content['body']['data'].encode("ASCII")).decode("utf-8")
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

    currDate = date.today()
    dayBefore = (currDate + timedelta(days = 1)).strftime("%y/%m/%d")
    weekBefore = (currDate - timedelta(days=60)).strftime("%y/%m/%d")
    message["Subject"] = f"Weekly Digest {daysBefore}~{weekBefore}"
    message.add_alternative(digest, subtype="html")
    #sendEmail(message)

def sendEmail(message):
    """ Takes EmailMessage and sends it to the receiver """
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    create_message = {'raw': encoded_message}
    send_message = msgs.send(userId = "me", body = create_message).execute()
    print("Email has been sent!")
    
if __name__ == '__main__':
    main()
