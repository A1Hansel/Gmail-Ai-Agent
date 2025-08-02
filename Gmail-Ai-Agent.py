#Gmail-Ai-Agent

import os
import datetime
import base64
import email
import requests
import pytz
from imapclient import IMAPClient
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from bs4 import BeautifulSoup
from openai import OpenAI
from email.mime.text import MIMEText
from email.utils import parseaddr
from email.utils import parsedate_to_datetime
from email.utils import formataddr
from dateutil import parser 
PhoneNum = os.getenv("PhoneNum") 

#       ~~~~~~~~~~~~~~~~~~~~~~   Settings   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MyName = 'Andrew'                  #Put your name here
EMAIL = 'a123hansel@gmail.com'
TextNotification = False                    #True if you want to recieve texts
PHONE_EMAIL = (f"{PhoneNum}@vtext.com")        #Only needed for text notifications  (Put your phone number where it says PhoneNum) and change the @vtext.com depending on your carrier
local_tz = pytz.timezone("America/Denver")        #Put in timezone

days_back = 1        #How many days of emails the agent will look back upon
TimeDelta = 2         #How many days ahead of the planned day should the AI try to reschedule
AvailableTimeCount = 3       #How many available times the AI will reccommend the user when scheduling


# --- Configurations ---
IMAP_SERVER = 'imap.gmail.com'
OPENAI_KEY = os.getenv("OPENAI_API_KEY1") #My own API key, you will have to create your own from OpenAI
Client = OpenAI(api_key=OPENAI_KEY)
SCOPES = ['https://mail.google.com/','https://www.googleapis.com/auth/calendar'] # Grants full access to Gmail and Calendar




REPLIED_MSGID_FILE = "replied_msgids.txt"
def get_oauth2_credentials(): #Generates and encrypts inputs for IMAP
    creds = None
    # token.json stores the user's access and refresh tokens.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If no valid creds, log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token_file:
            token_file.write(creds.to_json())
    return creds

def generate_xoauth2_string(email, access_token):
    auth_string = f'user={email}\1auth=Bearer {access_token}\1\1'
    return base64.b64encode(auth_string.encode()).decode()

#Calling login functions
creds = get_oauth2_credentials()
access_token = creds.token
auth_string = generate_xoauth2_string(EMAIL, access_token)


def load_replied_msgids():
    if not os.path.exists(REPLIED_MSGID_FILE):
        return set()
    with open(REPLIED_MSGID_FILE, "r") as f:
        return set(line.strip() for line in f.readlines())


def get_sender_name_and_received_time(msg):
   from_header = msg.get('From', '')
   date_header = msg.get('Date', '')

   name, addr = parseaddr(from_header)
   sender_name = name.strip() if name else addr.split('@')[0].replace('.', ' ').replace('_', ' ').title()

   try:
       received_time = parsedate_to_datetime(date_header)
   except Exception:
       received_time = None

   return [sender_name, received_time]

def save_replied_msgid(msgid):
    with open(REPLIED_MSGID_FILE, "a") as f:
        f.write(f"{msgid}\n")
        
        
# --- Functions for Google Calendar and Gmail ---

def get_calendar_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)
def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)


def find_gmail_id_from_header(gmail_service, message_id_header): 
    try:
        results = gmail_service.users().messages().list(
            userId='me',
            q=f'rfc822msgid:{message_id_header}'
        ).execute()
        messages = results.get('messages', [])
        if not messages:
            print(f"❌ No Gmail message found for Message-ID: {message_id_header}")
            return None
        return messages[0]['id']
    except Exception as e:
        print(f"Error finding Gmail ID for {message_id_header}: {e}")
        return None


def SearchCalendarForAvailableTime(DAY):
    available_slots = []

    # Loop through Days starting with tomorrow
    for i in range(TimeDelta+1):
        date = DAY + datetime.timedelta(days=i)
        start_of_day = datetime.datetime.combine(date, datetime.time(9, 0)).isoformat() + 'Z'
        end_of_day = datetime.datetime.combine(date, datetime.time(17, 0)).isoformat() + 'Z'

        events_result = calendar_service.events().list(
            calendarId='primary',
            timeMin=start_of_day,
            timeMax=end_of_day,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])

        busy_times = []
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            end = event['end'].get('dateTime', event['end'].get('date'))
            start_dt = parser.isoparse(start)
            end_dt = parser.isoparse(end)
            busy_times.append((start_dt, end_dt))


        # Check each 1-hour slot between 8am and 10pm
        for hour in range(8, 22):
            slot_start = datetime.datetime.combine(date, datetime.time(hour, 0))
            slot_end = slot_start + datetime.timedelta(hours=1)

            slot_start = slot_start.replace(tzinfo=datetime.timezone.utc)  ###Accounts for time zones
            slot_end = slot_end.replace(tzinfo=datetime.timezone.utc)

            if all(slot_end <= busy[0] or slot_start >= busy[1] for busy in busy_times):
                available_slots.append((slot_start, slot_end))

            if len(available_slots) == AvailableTimeCount: #Limits the availability response
               return available_slots
    return available_slots


def ConvertSlotsToLocalTime(available_slots):  
    converted_slots = []

    for start_utc, end_utc in available_slots:
        start_local = start_utc.astimezone(local_tz)
        end_local = end_utc.astimezone(local_tz)
        converted_slots.append((start_local, end_local))
    return converted_slots

def DoesTimeWork(UTCReqTime):
    utc_time = UTCReqTime
    end_time = utc_time + datetime.timedelta(hours=1)    
    
    # Query Google Calendar for events during this window
    events_result = calendar_service.events().list(
        calendarId='primary',
        timeMin=UTCReqTime.isoformat(),
        timeMax=end_time.isoformat(),
        singleEvents = True,
        orderBy='startTime'
    ).execute()
    events = events_result.get('items', [])
    if not events:
        print("✅ Time is available.")
        return True  # No conflicts found — the time works
    else:
        print("⛔ Time not available.")
        return False  # Conflicting events found — time is not free
    
    
    
def MarkCalendar(ConfirmedTime, subject, Name):
    end_time = ConfirmedTime + datetime.timedelta(hours=1)
    event = {
        'summary': f'AI: {subject} w/ {Name}',
        'description': 'This meeting was scheduled automatically based on availability.',
        'start': {
            'dateTime': ConfirmedTime.isoformat(),
            'timeZone': 'America/Denver',  
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'America/Denver',
        },
        'reminders': {
            'useDefault': True,
        },
    }

    try:
        calendar_service.events().insert(calendarId='primary', body=event).execute()
        print(f"📅 Event successfully created: {ConfirmedTime} to {end_time}")
    except Exception as e:
        print(f"❌ Failed to create calendar event: {e}")
    
    
# --- Connect to Email ---
def fetch_recent_emails(days_back):
    with IMAPClient(IMAP_SERVER, ssl=True) as client:
       client.oauth2_login(EMAIL, access_token)
       client.select_folder('INBOX')
       since_date = (datetime.datetime.now() - datetime.timedelta(days=days_back)).date()
       since_str = since_date.strftime('%d-%b-%Y')
       msgs = client.search(['SINCE', since_str])
       response = client.fetch(msgs, ['RFC822'])
       return [email.message_from_bytes(data[b'RFC822']) for data in response.values()]

# ---  text from email ---
def extract_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                return part.get_payload(decode=True).decode(errors='ignore')
            elif part.get_content_type() == 'text/html':
                html = part.get_payload(decode=True).decode(errors='ignore')
                return BeautifulSoup(html, 'html.parser').get_text()
    else:
        return msg.get_payload(decode=True).decode(errors='ignore')

def is_automated_email(subject, sender_email, body):
    automated_keywords = [
        "do not reply", "automated message", "no-reply", "noreply", 
        "auto-generated", "you are receiving this", "system notification"
    ]
    header = f"{subject} {sender_email}".lower()
    body = body.lower()
    
    for keyword in automated_keywords:
        if keyword in header or keyword in body:
            return "Yes"
    
    # Don't flag short human responses like "Yes, 14:00 tomorrow works"
    if len(body.split()) < 40 and "yes" in body:
        return "No"
    
    return "No"

# --- Classify email with OpenAI ---
def IntentToMeet(subject, sender, body):
    prompt = f"""
You are an AI assistant that reads emails sent to {MyName}.

Please determine the following:

   Does this email show intent to meet up with {MyName} sometime in the future? 
   This includes casual, polite, or direct requests, commands, or invitations to meet.
   Examples:
   - "Can you make it?"
   - "That works, see you then."
   - "Let's schedule a call."
   - "Come into work at 8:00 on Saturday."
   - "Please stop by on Friday afternoon."
   - "Join me on a call tomorrow at noon"
   - "We need to do ___ on Thursday"
   
Email details:
Email Subject: {subject}
From: {sender}
Body: {body[:2000]}

Respond with just one word: "Yes" or "No"
"""
    try:
        response = Client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=1,
        )

        result = response.choices[0].message.content.strip().lower()
        print(f"intent to meet = {result}")
        return result == "yes"
    except Exception as e:
        print(f"[IntentToMeet Error] {e}")
        return False #Fail-safe: don't act unless sure

#Finds what time The person wants to meet
def ExtractTime(subject, sender, body, reference_date_str, local_tz): ### Added local_tz part
    prompt = f"""
You are an AI assistant that reads emails sent to {MyName}.

Your job is to extract the **exact time and date** mentioned for any future meeting or event in the email.
If there is a time, it will be in {local_tz} time. You will need to output the date and time in the following format: ISO 8601 UTC (e.g., "2025-07-24T15:00:00Z").

Rules:
- Interpret informal phrases like "tomorrow at 3", "next Monday", "noon Tuesday" using the provided reference date.
- If no date/time is mentioned, reply with exactly: "None".
- If multiple times are mentioned, return the **earliest future one**.
- If the email only says something vague like "next week", try your best to estimate a time like "2025-08-28T12:00:00Z".

Reference date (i.e., when this email was received): {reference_date_str}
**ONLY RESPOND IN ISO 8601 UTC OR "NONE"**
Email Subject: {subject}
From: {sender}
Body:
{body[:2000]}
"""
    response = Client.chat.completions.create(
       model="gpt-3.5-turbo",
       messages=[{"role": "user", "content": prompt}],
       max_tokens = 40,
    )
    extracted = response.choices[0].message.content.strip()
    print(f"extracted time = {extracted}")
    return extracted
    
    
#~~~~~~~~~~~~ Types of Responses to Give ~~~~~~~~~~~~~~~~~~~~~~~
    
def TimeDoesNotWorkResponse(AvailableSlots, Name, body, ReqTime):
    prompt = f"""
You are an AI assistant that writes emails on {MyName}'s behalf.

Your job is to tell {Name}, that {MyName} is busy at {ReqTime}, but that he is available at the times: {AvailableSlots}. 
Reference the content of {Name}'s email below when crafting your response. Be clear, casual, and concise.
Also, make it known that you are {MyName}'s AI assisstant
Email from {Name}:
{body[:2000]}
"""
    response = Client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )
    print('AI response created')
    return response.choices[0].message.content.strip()


def SuggestANewTimeResponse(AvailableSlots, Name, body):
    prompt = f"""
You are an AI assistant that writes emails on {MyName}'s behalf.

Your job is to suggest scheduling times to {Name}, based on {MyName}'s availability: {AvailableSlots}.

Your job is to tell {Name}, that {MyName} is busy, but that he is available at the times: {AvailableSlots}. 
Reference the content of {Name}'s email below when crafting your response. Be clear, casual, and concise.
Also, make it known that you are {MyName}'s AI assisstant
Email from {Name}:
{body[:2000]}
"""
    response = Client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )
    print('AI response created')
    return response.choices[0].message.content.strip()


def ThatTimeWorksResponse(ConfirmedTime, Name, body):    
    prompt = f"""
 You are an AI assistant that writes emails on {MyName}'s behalf.

 Your job is to confirm the meeting time of {ConfirmedTime} to {Name}.
 Reference the content of {Name}'s email below when crafting your response. Be clear, casual, concise, and excited.
 Also, make it known that you are {MyName}'s AI assisstant
 
 Email from {Name}:
 {body[:2000]}
 """
    response = Client.chat.completions.create(
         model="gpt-3.5-turbo",
         messages=[{"role": "user", "content": prompt}]
     )
    print('AI response created')
    return response.choices[0].message.content.strip()



# --- Sends email response ---
def send_response(sender_email, ai_reply, access_token, og_message_id, thread_id, subject):
    msg = MIMEText(ai_reply)
    msg['To'] = sender_email
    msg['From'] = formataddr(('AI Agent', EMAIL))
    msg['Subject'] = f"Re: {subject}"
    msg['In-Reply-To'] = og_message_id
    msg['References'] = og_message_id

    raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    payload = {
        'raw': raw_msg,
        'threadId': thread_id  # This keeps it in the original thread
    }

    response = requests.post(
        'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        },
        json=payload
    )

    if response.status_code == 200:
        print("✅ Response email sent.")
    else:
        print(f"❌ Failed to send email: {response.status_code} – {response.text}")


def send_text_notification(PHONE_EMAIL, Name, ConfirmedTime, subject): 
       msg = MIMEText(f" {subject} with {Name} at {ConfirmedTime}")
       msg['To'] = PHONE_EMAIL
       msg['From'] = formataddr(('AI Agent', EMAIL))
       raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
       payload = {
           'raw': raw_msg
       }
       response = requests.post(
         'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
         headers={
            'Authorization': f'Bearer {access_token}',
             'Content-Type': 'application/json'
            },
            json=payload
       )

       if response.status_code == 200:
         print("✅ Text message sent via email-to-SMS gateway.")
       else:
         print(f"❌ Failed to send message: {response.status_code} – {response.text}") 

     
             
# --- Main agent loop ---
def run_agent():
    global  calendar_service, gmail_service, og_message_id, thread_id #makes all functions outside of agent() able to use these variables
    calendar_service = get_calendar_service() # For Calendar API calls
    gmail_service = get_gmail_service()        # For Gmail API calls
    replied_msgids = load_replied_msgids()

    with IMAPClient(IMAP_SERVER, ssl=True) as client:
        client.oauth2_login(EMAIL, access_token)
        client.select_folder('INBOX')
        since_date = (datetime.datetime.now() - datetime.timedelta(days=days_back)).date()
        since_str = since_date.strftime('%d-%b-%Y')
        msgs = client.search(['SINCE', since_str])
        response = client.fetch(msgs, ['RFC822'])

        for msgid, data in response.items():
            raw_email = data[b'RFC822']
            msg = email.message_from_bytes(raw_email)
         
            message_id = msg.get("Message-ID")
            in_reply_to = msg.get('In-Reply-To')
            
            
            if not message_id: #Skip if there is no message ID
                print("⚠️ No Message-ID found. Skipping.")
                continue
                
            # Skip if already replied
            if message_id in replied_msgids:
                print(f"🟡 Already replied to message ID {message_id}, skipping.")
                continue
            
            # Skip self-sent messages
            sender_name, sender_email = parseaddr(msg.get('From', ''))
            if sender_email.lower() == EMAIL.lower() or sender_email.lower() == 'ai agent': 
                print(f"⚠️ Skipping self-sent email from {sender_email}")
                continue

            subject = msg['subject']
            body = extract_body(msg)
            
            #Skips Automated Messages
            is_auto = is_automated_email(subject, sender_email, body)
            if is_auto.lower().strip() == "yes":
                print("🤖 Automated message detected. Skipping response.")
                continue
            
            #Gets Thread-ID
            gmail_id = find_gmail_id_from_header(gmail_service, message_id)
            if not gmail_id:
                continue  # Skip this email if no matching Gmail message found
   
            msg_metadata = gmail_service.users().messages().get(
                userId='me',
                id=gmail_id,
                format='metadata'
            ).execute()
                      
            #Gmail and Thread ID's
            thread_id = msg_metadata['threadId']
            
            og_message_id = next(
               (h['value'] for h in msg_metadata['payload']['headers'] if h['name'].lower() == 'message-id'),
               None
            )
            for header in msg_metadata['payload']['headers']:
                if header['name'].lower() == 'message-id':
                    og_message_id = header['value']
                    
            msg['In-Reply-To'] = og_message_id
            msg['References'] = og_message_id
            subject = msg['subject']       
            
            NameDate = get_sender_name_and_received_time(msg) # Determines the name of the user, and when the email was received
            Name = NameDate[0]
            #EmailTime = NameDate[1] ---> maybe use later###
            Now = datetime.datetime.now(local_tz).isoformat()
            
            #Checks if there is an intent to meet
            if not IntentToMeet(subject, sender_email, body): #If IntentToMeet is False:
                print(f"\n📩 Subject: {subject}")
                print(f"There is no intent to meet with {MyName}")
                continue
            
            #Finds a requested time if there is one
            ReqTimeStr = ExtractTime(subject, sender_email, body, Now, local_tz) #(in UTC)
         
            
            today = datetime.datetime.utcnow().replace(hour=8, minute=0, second=0, microsecond=0)
            tomorrow = today + datetime.timedelta(days=1)
            
            if ReqTimeStr == "None":
                AvailableSlots = SearchCalendarForAvailableTime(tomorrow) 
                ConvertedSlots = ConvertSlotsToLocalTime(AvailableSlots)                         ###
                 
                ai_reply = SuggestANewTimeResponse(ConvertedSlots, Name, body)
                send_response(sender_email, ai_reply, access_token, og_message_id, thread_id, subject)
                print(f"\n📩 Subject: {subject}")
                print(f'Suggested meeting times to {Name}')
                continue
            else:
                TimeObj = datetime.datetime.strptime(ReqTimeStr, "%Y-%m-%dT%H:%M:%SZ") 
                TimeObj = TimeObj.replace(tzinfo=pytz.utc)  # Converted to UTC more cleanly
                ReqTime = TimeObj.astimezone(local_tz) #ReqTime is now on local time
                
                print(f"Requested time of {ReqTime} detected")
                if DoesTimeWork(TimeObj) == True:      #Check if ReqTime works
                    ConfirmedTime = ReqTime
                    MarkCalendar(TimeObj, subject, Name)
                    ai_reply = ThatTimeWorksResponse(ConfirmedTime, Name, body)
                    send_response(sender_email, ai_reply, access_token, og_message_id, thread_id, subject)
                    
                    if TextNotification == True:     #Sends Textmessage
                        send_text_notification(PHONE_EMAIL, Name, ConfirmedTime, subject)
                    
                else:
                    AvailableSlots = SearchCalendarForAvailableTime(tomorrow) 
                    ConvertedSlots = ConvertSlotsToLocalTime(AvailableSlots)  
                    
                    ai_reply = TimeDoesNotWorkResponse(ConvertedSlots, Name, body, ReqTime)
                    send_response(sender_email, ai_reply, access_token, og_message_id, thread_id, subject)
                    
         # Save message IDs to avoid duplicate replies
        if message_id:
                    save_replied_msgid(message_id)
        if in_reply_to:
                    save_replied_msgid(in_reply_to)
            
        print('Agent Executed')
        

try:
    run_agent()
except KeyboardInterrupt: #Allows the program to be interuppted
    print("🛑 Agent execution interrupted by user.")




