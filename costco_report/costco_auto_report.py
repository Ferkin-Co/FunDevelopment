from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime as dt
import pandas as pd
import traceback
import base64
import zipfile
import os
import io

"""
REGION, LOCATION, DEPARTIME, ACTIVITY NAME, LAST NAME, FIRST NAME, HIRED DATE, JOB CODE, JOB TITLES, STATUS, GRADE, 
COMPLETION DATE, COMMENTS

ignore = Region, Location ID, Hired Date, Job Code, Job Title, Status, Grade, Comments

- No duplicate names
- Certificates COMPLETION DATE onto same row as name (All certs)
    - Ex: Costco U Cert = MM/DD/YY, etc.
- If missing any certs, Missing = True

WHEN FIRST RUNNING PROGRAM, YOU WILL NEED TO AUTHENTICATE USER
- If you need to reset your credentials/token, delete "token.json"

"""
# CONSTANTS
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
SEND_TO = "gjitmetta@gmail.com"
# SEND_TO = "palaszewskisteven@gmail.com"
TODAY_DATE = dt.datetime.today().strftime("%Y-%m-%d")
TOKEN_FILE = "token.json"
EMAIL_SUBJECT = "Background Report Job Email Notification"

# NAME OF CSV TO EDIT FROM ZIP
CSV_NAME = "Activity Completions (CSV).csv"

#COLUMNS AND DICTIONARY MAPPING
drop_cols = ["Region", "Location ID", "Hired Date", "Job Code", "Job Title", "Status", "Grade", "Comments"]
gas_station_tr_headers = ["Dept", "Employee", "Costco U Cert", "Class C p1", "Class C p2", "JHA", "Facility Walk",
                          "Missing", "Activity Name"]
certs_cols = ["Costco U Cert", "Class C p1", "Class C p2"]
gas_cert_mapping = {"Gas Station Certification - US": "Costco U Cert",
                    "Gas: Part 1: Class C Operator Training and Exam by USTtraining.com": "Class C p1",
                    "Gas: Part 2: Class C UST Operator Training - Facility Specific Worksheet Acknowledgment": "Class C p2"}


def gmail_authenticate():  # Create GMAIl credentials for authentication
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
    gmail_service = build("gmail", "v1", credentials=creds)
    return gmail_service


def get_gmail_service():  # get credentials
    service = gmail_authenticate()
    return service


def get_email():  # Get email containing specified subject -- grabs most recent
    results = service.users().messages().list(userId='me', q=f'subject:"{EMAIL_SUBJECT}"',
                                              maxResults=1).execute()
    messages = results.get('messages', [])

    if messages:
        return messages[0]
    return


def get_email_content(message_id):  # get data from email (subject, body, etc)
    email_message = message_id
    if email_message:
        email_id = email_message['id']
        print(f"Downloading attachment from email message: Background Report Job Email Notification")

        source_dataframe = get_attachments(email_id)
        return source_dataframe


def get_attachments(message_id):  # get report attachment from email data and save in-memory
    message = service.users().messages().get(userId="me", id=message_id).execute()

    parts = message['payload'].get('parts', [])
    for part in parts:
        if part['filename']:
            file_name = part['filename']
            if 'attachmentId' in part['body']:
                attachment_id = part['body']['attachmentId']
                attachment = service.users().messages().attachments().get(userId="me", messageId=message_id,
                                                                          id=attachment_id).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                print(f"Unzipping {part['filename']}...")

                if file_name.lower().endswith('.zip'):
                    with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zipped_file:
                        with zipped_file.open(CSV_NAME) as file:
                            dataframe = pd.read_csv(file, skiprows=3)
                            return dataframe


def create_activity_df():  # Create dataframe from unzipped csv file in-memory
    activity_comp_df = source_df

    activity_comp_df.drop(columns=drop_cols, inplace=True)
    activity_comp_df["Completion Date"] = pd.to_datetime(activity_comp_df["Completion Date"]).dt.strftime('%m/%d/%y')
    activity_comp_df["Department"] = activity_comp_df["Department"].str[5:]

    return activity_comp_df


def create_gas_df(activity_df):  # create formatted dataframe
    gas_train_df = pd.DataFrame(columns=gas_station_tr_headers)
    activity_comp_df = activity_df
    gas_train_df["Employee"] = activity_comp_df[["First Name", "Last Name"]].agg(' '.join, axis=1).str.title()
    activity_comp_df["Employee"] = gas_train_df["Employee"]

    gas_train_df["Dept"] = activity_comp_df["Department"]
    gas_train_df["Completion Date"] = activity_comp_df["Completion Date"]

    return gas_train_df


def create_result_df(gas_df):  # Create result dataframe
    gas_train_df = gas_df

    result_df = gas_train_df.copy()
    result_df.sort_values(by=["Employee", "Completion Date"], ascending=[True, False], inplace=True)
    result_df.reset_index(drop=True, inplace=True)

    return result_df


def format_result_df(activity_df, res_df):  # Format dataframe
    activity_comp_df = activity_df
    result_df = res_df

    for index, row in activity_comp_df.iterrows():
        # STORE ROW DATA INTO VARIABLES PER ITERATION
        name = row["Employee"]
        certs = row["Activity Name"]
        completion = row["Completion Date"]

        # GET VALUE OF GAS_CERT_MAPPING IF KEY == CERTS
        map_cols = gas_cert_mapping.get(certs)

        # IF MAP_COLS NOT EMPTY
        if map_cols:
            # AT LOCATION OF EMPLOYEE ROW WHERE DATA IN EMPLOYEE COLUMN AND CERTS COLUMN MATCHES
            # SET CERT COLUMN OF THAT ROW == COMPLETION DATE
            result_df.loc[result_df["Employee"] == name, map_cols] = completion

    # DROP UNUSED COLUMNS AND DEDUPE
    result_df.drop(columns=["Activity Name", "Completion Date"], axis=1, inplace=True)
    result_df = result_df.drop_duplicates(subset="Employee", keep='first')

    # ITERATE THROUGH ROWS IN RESULT DF
    for index, row in result_df.iterrows():
        # IF ANY COLUMN IN A ROW THAT IS IN CERTS_COLS IS NULL, MISSING_CERT == FALSE, ELSE TRUE
        missing_cert = any(pd.isna(row[col]) for col in certs_cols)

        # RETURN MISSING_CERT VALUE TO MISSING COLUMN IN ROW AT INDEX
        result_df.at[index, "Missing"] = missing_cert

    # RESET INDEX TO START AT 0
    result_df.reset_index(drop=True, inplace=True)

    return result_df


def export_log(res_df):  # export resulting dataframe as csv and send to email
    result_df = res_df
    buffer = io.BytesIO()

    result_df.to_csv(buffer, index=False)
    buffer.seek(0)

    message = MIMEMultipart()
    message['to'] = SEND_TO
    message['subject'] = f"Costco Activity-Training Log {TODAY_DATE}"
    message.attach(MIMEText("See Attached File", "plain"))

    # create attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(buffer.read())

    # encode attachment in base64
    encoders.encode_base64(part)

    # Add header to specify name of attachment
    part.add_header("Content-Disposition", f"attachment; filename=Costco Activity-Training Log {TODAY_DATE}.csv")
    message.attach(part)

    # Encode message in base64 format
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId='me', body={"raw" : raw_message}).execute()

    return


def execute_log():  # execute creation of log and sending it
    activity_df = create_activity_df()
    gas_df = create_gas_df(activity_df)
    result_df = format_result_df(activity_df, create_result_df(gas_df))
    export_log(result_df)
    print(f"Program Completed -- Report sent to {SEND_TO}")
    return


try:  # execute final steps
    service = get_gmail_service()
    message_id = get_email()
    source_df = get_email_content(message_id)
    execute_log()
except Exception as e:
    traceback_info = traceback.format_exc()
    print(f"An error has occured: {e}\n\nTraceback Info: {traceback_info}")
