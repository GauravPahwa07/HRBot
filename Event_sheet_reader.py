import gspread
import ssl
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from slack_sdk import WebClient

def read_google_sheet():
    # Path to the credentials JSON file you downloaded
    credentials_file = r'sheet_api.json'

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

    # Get the current date (ignoring the year)
    current_date = datetime.now().date().replace(year=datetime.now().year)

    # Set the SSL certificate location
    ssl._create_default_https_context = ssl._create_unverified_context

    # Initialize the Slack client
    slack_token = 'xoxb-7556033825-5569483529268-DFwTVksCMy36Fw06PyZn9msy'
    slack_client = WebClient(token=slack_token)

    # Process the data
    for row in data:
        # Access individual fields
        name = row['Name']
        email = row['Email']
        event_date = datetime.strptime(row['Date'], "%d %B %Y").date()
        event_type = row['Event']
        employee_status = row['Employee Status']

        # Check if the employee is not an ex-employee and the event date (ignoring the year) is today
        if (
            employee_status != 'ex'
            and event_date.day == current_date.day
            and event_date.month == current_date.month
        ):
            # Customize the message based on the event type
            if event_type == 'Birthday':
                message = f"Happy birthday, {name}! May your day be filled with joy and laughter. 🎉🎂"
            elif event_type == 'DOJ':
                message = f"Congratulations, {name}, on your work anniversary! Thank you for your dedication and hard work. 🎉🎊"
            elif event_type == 'Anniversary':
                message = f"Happy anniversary, {name}! Wishing you many more years of love and happiness together. 💑💍"

            # Post the message to the Slack channel
            response = slack_client.chat_postMessage(channel='C05JE13BDTJ', text=message)

            # Handle API errors
            if response["ok"]:
                print(f"Message posted for {name} in Slack.")
            else:
                print(f"Error posting message for {name} in Slack: {response['error']}")

# Call the function to read the Google Sheet and identify events today
read_google_sheet()
