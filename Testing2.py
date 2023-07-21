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
    slack_token = 'xoxb-7556033825-5569483529268-3Kw4tClSI3RKd7dhygcxTbZt'
    slack_client = WebClient(token=slack_token)

    # Define the function to upload an image to Slack
    def upload_image_to_slack(image_url, token):
        try:
            image_response = requests.get(image_url)
            if image_response.status_code == 200:
                image_content = image_response.content

                # Use the Slack WebClient to upload the file
                image_upload_response = slack_client.files_upload(
                    channels="C05JE13BDTJ",
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
        event_date = datetime.strptime(row['Date'], "%d %B %Y").date()
        event_type = row['Event']
        employee_status = row['Employee Status']

        # Check if the employee is not an ex-employee and the event date (ignoring the year) is today
        if (
            employee_status != 'ex'
            and event_date.day == current_date.day
            and event_date.month == current_date.month
        ):
            # Check if the birthday message has been posted already
            if not row.get('Birthday_Message_Posted') and event_type == 'Birthday':
                # Mark the birthday message as posted in the data
                row['Birthday_Message_Posted'] = True

                # Customize the birthday message
                message = f"Happy birthday, {name}! May your day be filled with joy and laughter. 🎉🎂"

                try:
                    response = slack_client.chat_postMessage(channel='C05JE13BDTJ', text=message)
                    # Handle API errors
                    if response["ok"]:
                        print(f"Message posted for {name} in Slack.")
                    else:
                        print(f"Error posting message for {name} in Slack: {response['error']}")
                except SlackApiError as e:
                    print(f"Error posting message for {name} in Slack: {e.response['error']}")

                # Add a slight delay before uploading the image (optional, to ensure the message is posted first)
                time.sleep(1)

                image_url = 'https://drive.google.com/uc?export=download&id=1EZN2zNH9CFS23KbfZYW4W1U4_JLQ4Dcy'  # Replace with the actual URL of the image in Google Drive
                if image_url:
                    # Upload the image to Slack
                    image_private_url = upload_image_to_slack(image_url, slack_token)
                    if image_private_url:
                        print(f"Image uploaded to Slack for {name}.")
                    else:
                        print(f"Error uploading image to Slack for {name}.")

            elif event_type == 'DOJ':
                # Check if the work anniversary message has been posted already
                if not row.get('Work_Anniversary_Message_Posted'):
                    # Mark the work anniversary message as posted in the data
                    row['Work_Anniversary_Message_Posted'] = True

                    message = f"Congratulations, {name}, on your work anniversary! Thank you for your dedication and hard work. 🎉🎊"
                    try:
                        response = slack_client.chat_postMessage(channel='C05JE13BDTJ', text=message)
                        # Handle API errors
                        if response["ok"]:
                            print(f"Message posted for {name} in Slack.")
                        else:
                            print(f"Error posting message for {name} in Slack: {response['error']}")
                    except SlackApiError as e:
                        print(f"Error posting message for {name} in Slack: {e.response['error']}")

    # Update the data back to the Google Sheet to mark the birthday and work anniversary messages as posted
    sheet.update('A2', data)

# Call the function to read the Google Sheet and identify events today
read_google_sheet()
