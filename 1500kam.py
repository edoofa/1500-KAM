import os
import pickle
import pandas as pd
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from googleapiclient.discovery import build
import json
import re

# Define scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Authenticate and create the service
def authenticate():
    creds = None
    if os.path.exists('token.json'):
        with open('token.json', 'r') as token:
            creds_info = json.load(token)
            creds = Credentials.from_authorized_user_info(creds_info, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

# Authentication
creds = authenticate()
client = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)

# Function to get files created this month containing "Sales Report" in the name
def get_sales_report_files():
    today = datetime.now()
    first_day_of_month = today.replace(day=1).isoformat() + 'Z'
    query = f"name contains 'Sales Report' and mimeType='application/vnd.google-apps.spreadsheet' and createdTime >= '{first_day_of_month}'"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    file_urls = []
    for item in items:
        file_id = item['id']
        file_urls.append(f"https://docs.google.com/spreadsheets/d/{file_id}")
    return file_urls

# Get daily report URLs
daily_report_urls = get_sales_report_files()

# Function to get data from a Google Sheet
def get_sheet_data(sheet_url, sheet_name):
    sheet = client.open_by_url(sheet_url).worksheet(sheet_name)
    return pd.DataFrame(sheet.get_all_records())

# Read the Priority cases sheet
priority_cases_url = "https://docs.google.com/spreadsheets/d/1UBm4B-THlm-4EU_sCfNIudCYG0yBWk_wfrNWGeSFu78/edit#gid=0"
priority_cases_sheet = client.open_by_url(priority_cases_url).worksheet('KAM')
priority_cases_df = pd.DataFrame(priority_cases_sheet.get_all_records())

# Read the Mastersheet
mastersheet_url = "https://docs.google.com/spreadsheets/d/1zWkj6dgfEFBlN6eQne5kF6JqMrnNl7ybQIcY6D4qP0Y/edit#gid=2085346559"
mastersheet = client.open_by_url(mastersheet_url).worksheet('Mastersheet')
mastersheet_df = pd.DataFrame(mastersheet.get_all_records())

# Remove any pre-existing incorrect columns
incorrect_columns = ['Date COMSBS', 'Date COMSBE']
priority_cases_df = priority_cases_df.drop(columns=[col for col in incorrect_columns if col in priority_cases_df.columns])

# Debug: Print column names to verify
print("Priority cases columns:", priority_cases_df.columns.tolist())
print("Mastersheet columns:", mastersheet_df.columns.tolist())

# Ensure the columns are present in the Priority cases DataFrame
if 'KAM Status' not in priority_cases_df.columns:
    priority_cases_df['KAM Status'] = ''
if 'Intake' not in priority_cases_df.columns:
    priority_cases_df['Intake'] = ''
if 'Admission Officer' not in priority_cases_df.columns:
    priority_cases_df['Admission Officer'] = ''

# Map Group Name to additional columns from Mastersheet
for index, row in priority_cases_df.iterrows():
    group_name = row['KAM Group Name']
    master_match = mastersheet_df[mastersheet_df['Admission Group Name'].str.contains(re.escape(group_name), case=False, na=False)]
    if not master_match.empty:
        master_data = master_match.iloc[0]
        priority_cases_df.at[index, 'KAM Status'] = master_data['KAM Status']
        priority_cases_df.at[index, 'Intake'] = master_data['Intake']
        priority_cases_df.at[index, 'Admission Officer'] = master_data['Admission Officer']

# Function to update Priority cases sheet
def update_priority_cases(priority_cases_df, daily_report_df, date_str):
    # Debug: Print column names of daily report
    print("Daily report columns:", daily_report_df.columns.tolist())
    
    for index, row in priority_cases_df.iterrows():
        student_name = row['KAM Group Name']
        # Match the student_name partially with chat file names
        match = daily_report_df[daily_report_df['Chat File Name'].str.contains(re.escape(student_name), case=False, na=False)]
        if not match.empty:
            student_data = match.iloc[0]
            comsbs = student_data['Count_of_message_send_by_student']
            comsbe = student_data['Count_of_message_send_by_employee']
            # Ensure we start putting values from Column E
            priority_cases_df.at[index, f'{date_str} COMSBS'] = comsbs
            priority_cases_df.at[index, f'{date_str} COMSBE'] = comsbe
            print(f"Match found: {student_name} -> {student_data['Chat File Name']}, Date: {date_str}, COMSBS: {comsbs}, COMSBE: {comsbe}")
        else:
            priority_cases_df.at[index, f'{date_str} COMSBS'] = 0
            priority_cases_df.at[index, f'{date_str} COMSBE'] = 0
            print(f"No match found for: {student_name} on Date: {date_str}")

# Process each Daily Report sheet
for url in daily_report_urls:
    print(f"Processing sheet: {url}")
    daily_report_df = get_sheet_data(url, 'Input File')
    
    # Extract the date from the first row and first column
    date_str = daily_report_df.iloc[0, 0]
    print(f"Date extracted: {date_str}")
    
    update_priority_cases(priority_cases_df, daily_report_df, date_str)

# Sort the date columns
date_columns = [col for col in priority_cases_df.columns if re.match(r'\d{4}-\d{2}-\d{2} COMSBS|\d{4}-\d{2}-\d{2} COMSBE', col)]
date_columns_sorted = sorted(date_columns, key=lambda x: datetime.strptime(x.split()[0], '%Y-%m-%d'))

# Calculate the sum of COMSBS and COMSBE for each row
priority_cases_df['Total STU'] = priority_cases_df[[col for col in date_columns_sorted if 'COMSBS' in col]].sum(axis=1)
priority_cases_df['Total EMP'] = priority_cases_df[[col for col in date_columns_sorted if 'COMSBE' in col]].sum(axis=1)

# Calculate the count of COMSBS and COMSBE > 0 for each row
priority_cases_df['Days STU'] = priority_cases_df[[col for col in date_columns_sorted if 'COMSBS' in col]].gt(0).sum(axis=1)
priority_cases_df['Days EMP'] = priority_cases_df[[col for col in date_columns_sorted if 'COMSBE' in col]].gt(0).sum(axis=1)

# Reorganize DataFrame
non_date_columns = [col for col in priority_cases_df.columns if col not in date_columns]
sorted_columns = ['KAM Group Name', 'KAM Status', 'Intake', 'Admission Officer', 'Total STU', 'Total EMP', 'Days STU', 'Days EMP'] + date_columns_sorted
priority_cases_df = priority_cases_df[sorted_columns]

# Update date column headers
new_headers = []
for col in priority_cases_df.columns:
    if re.match(r'\d{4}-\d{2}-\d{2}', col):
        date_part = datetime.strptime(col.split()[0], '%Y-%m-%d').strftime('%d %b %Y')
        if 'COMSBS' in col:
            new_headers.append(f"{date_part} STU")
        elif 'COMSBE' in col:
            new_headers.append(f"{date_part} EMP")
    else:
        new_headers.append(col)

priority_cases_df.columns = new_headers

# Prepare data for update
data_for_update = [priority_cases_df.columns.values.tolist()] + priority_cases_df.values.tolist()

# Update the Priority cases sheet with the new data, starting from Column A
priority_cases_sheet.update('A1', data_for_update)
