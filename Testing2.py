import gspread
import ssl
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import requests
import time

def read_google_sheet():
    # Path to the credentials JSON file you downloaded
    credentials_file = 'sheet_api.json'

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
    slack_token = 'xoxb-7556033825-5569483529268-IhMYBKc5NjtdHtSOkfM6xptC'
    slack_client = WebClient(token=slack_token)

    # Define the function to upload an image to Slack
    def upload_image_to_slack(image_url, token):
        try:
            image_response = requests.get(image_url)
            if image_response.status_code == 200:
                image_content = image_response.content

                # Use the Slack WebClient to upload the file
                image_upload_response = slack_client.files_upload(
                    channels="C02JXF77LKH",
                    initial_comment="",  # Empty initial comment
                    file=image_content,  # Use the "file" parameter directly to specify the image content
                    filename="birthday_image.png",
                    filetype="png"  # Specify the file type as "png"
                )
                return image_upload_response
            else:
                print(f"Error downloading image from URL: {image_response.status_code}")
        except SlackApiError as e:
            print(f"Error uploading image to Slack: {e.response['error']}")
        except Exception as ex:
            print(f"Error: {ex}")

    # Process the data
    for row in data:
        # Access individual fields
        name = row['Name']
        email = row['Email']
        dob_date_str = row['DOB']  # Assuming DOB format is "19 July 1994"
        dob_date = datetime.strptime(dob_date_str, "%d %B %Y").date().replace(year=datetime.now().year)
        doj_date_str = row['DOJ']  # Assuming DOJ format is "19 July 2023"
        doj_date = datetime.strptime(doj_date_str, "%d %B %Y").date().replace(year=datetime.now().year)
        employee_status = row['Employee Status']

        # Check if the employee is not an ex-employee and today is their birthday
        if (
            employee_status != 'ex'
            and dob_date.day == current_date.day
            and dob_date.month == current_date.month
        ):
            # Customize the birthday message
            message = f"We wish you a very Happy birthday @{name} ! May this special day bring you joy, laughter, and countless wonderful memories. Enjoy your day to the fullest!. 🎉🎂"

            try:
                response = slack_client.chat_postMessage(channel='C02JXF77LKH', text=message)
                # Handle API errors
                if response["ok"]:
                    print(f"Message posted for {name} in Slack (Birthday).")
                else:
                    print(f"Error posting message for {name} in Slack (Birthday): {response['error']}")
            except SlackApiError as e:
                print(f"Error posting message for {name} in Slack (Birthday): {e.response['error']}")

            # Add a slight delay before uploading the image (optional, to ensure the message is posted first)
            time.sleep(1)

            image_url = 'https://drive.google.com/uc?export=download&id=1EZN2zNH9CFS23KbfZYW4W1U4_JLQ4Dcy'  # Replace with the actual URL of the image in Google Drive
            if image_url:
                # Upload the image to Slack
                image_private_url = upload_image_to_slack(image_url, slack_token)
                if image_private_url:
                    print(f"Image uploaded to Slack for {name} (Birthday).")
                else:
                    print(f"Error uploading image to Slack for {name} (Birthday).")

        # Check if today is an employee's work anniversary
        if (
            employee_status != 'ex'
            and doj_date.day == current_date.day
            and doj_date.month == current_date.month
        ):
            # Customize the work anniversary message
            message = f"Congratulations, {name} on reaching another milestone in your career! Your dedication, commitment, and hard work inspire us all. Here's to many more successful years together !!! 🎉🎊"

            try:
                response = slack_client.chat_postMessage(channel='C02JXF77LKH', text=message)
                # Handle API errors
                if response["ok"]:
                    print(f"Message posted for {name} in Slack (Work Anniversary).")
                else:
                    print(f"Error posting message for {name} in Slack (Work Anniversary): {response['error']}")
            except SlackApiError as e:
                print(f"Error posting message for {name} in Slack (Work Anniversary): {e.response['error']}")

    print("All data processed.")

# Call the function to read the Google Sheet and identify events today
read_google_sheet()
