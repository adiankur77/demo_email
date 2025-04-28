import requests
import json

# Replace with your application details
TENANT_ID = "f44408bd-674c-490f-8e70-c0b65bc8f4cb"  # From Entra ID app registration overview
CLIENT_ID = "83a5bba6-b181-4230-a3cd-b5bccd44e61f"  # From Entra ID app registration overview
CLIENT_SECRET = "K528Q~DG8VEGY6JurZf98WpwkcFoErdknBQbicLo"  # From Entra ID app registration -> Certificates & secrets
MAILBOX_TO_SEND_FROM = "adityaankur55@outlook.com"  # The email address that will send the email

# Email details
TO_RECIPIENTS = [{"EmailAddress": {"Address": "recipient@example.com"}}]
SUBJECT = "Test Email from Microsoft Graph API"
BODY = {"ContentType": "Text", "Content": "This is a test email sent using the Microsoft Graph API and Python!"}

# Microsoft Graph API endpoint for sending emails
SEND_EMAIL_URL = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_TO_SEND_FROM}/sendMail"

# OAuth 2.0 token endpoint
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

def get_access_token():
    """Gets an access token using the client credentials grant flow."""
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(TOKEN_URL, data=data)
    response.raise_for_status()  # Raise an exception for bad status codes
    return response.json()['access_token']

def send_email(access_token, to_recipients, subject, body):
    """Sends an email using the Microsoft Graph API."""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    payload = {
        'message': {
            'subject': subject,
            'body': body,
            'toRecipients': to_recipients
        },
        'saveToSentItems': 'true'  # Optional: Save the email in the "Sent Items" folder
    }
    print("Request body ready")
    response = requests.post(SEND_EMAIL_URL, headers=headers, data=json.dumps(payload))
    print(f"Full response: {response.text}") 
    response.raise_for_status()  # Raise an exception for bad status codes
    print("Email sent successfully!")

if __name__ == "__main__":
    try:
        access_token = get_access_token()
        send_email(access_token, TO_RECIPIENTS, SUBJECT, BODY)
    except requests.exceptions.RequestException as e:
        print(f"Error during API request: {e}")
    except KeyError as e:
        print(f"Error parsing JSON response: Missing key - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
