import gspread
import ssl
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from slack_sdk import WebClient

def read_google_sheet():
    # Path to the credentials JSON file you downloaded
    credentials_file = r'C:\Users\gaura\Downloads\well-wisher-bot-9dc5af4bea3b.json'

    # Authenticate with the Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(credentials)

    # ID or URL of the Google Sheet
    sheet_id = '1lw1kRGhE0317oC35VHuMg3LB5uhB6IR6XrGzEIpVsJM'

    # Open the Google Sheet
    sheet = client.open_by_key(sheet_id).sheet1

    # Read the data from the sheet
    data = sheet.get_all_records()

    # Get the current date
    current_date = datetime.now().date()

    # Set the SSL certificate location
    ssl._create_default_https_context = ssl._create_unverified_context

    # Initialize the Slack client
    slack_token = 'xoxb-7556033825-5569483529268-R5yWhnNQWUBHr80OZMhJOG88'
    slack_client = WebClient(token=slack_token)

    # Process the data
    for row in data:
        # Access individual fields
        name = row['Name']
        email = row['Email']
        birthday = datetime.strptime(row['Birthday'], "%d %B %Y").date()
        department = row['Department']

        # Check if the birthday is today
        if birthday == current_date:
            # Post the birthday wish to the Slack channel
            message = f"Happy birthday, {name}!"
            response = slack_client.chat_postMessage(channel='C03A6JM3CU8', text=message)

            # Handle API errors
            if response["ok"]:
                print(f"Birthday wish posted for {name} in Slack.")
            else:
                print(f"Error posting birthday wish for {name} in Slack: {response['error']}")

# Call the function to read the Google Sheet and identify employees with a birthday today
read_google_sheet()
