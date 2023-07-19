import requests
import os
from getpass import getpass
from datetime import datetime
import json
import shutil
from bs4 import BeautifulSoup

def get_access_token(client_id: str, client_secret: str, tenant_id: str) -> str:
    """Obtain the access token using client credentials flow.

    Args:
        client_id (str): Client ID (Application ID) of the registered application.
        client_secret (str): Client Secret (Application Secret) of the registered application.
        tenant_id (str): Tenant ID associated with the registered application.

    Returns:
        str: Access token.
    """
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(url, data=data)
    response.raise_for_status()
    access_token = response.json()["access_token"]
    return access_token

def extract_text_from_html(html):
    """Extract text content from HTML."""
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator=' ')
    return text.strip()


def download_office365_attachments(access_token: str, start_date: datetime.date, start_time: datetime.time) -> tuple[dict, dict]:
    """Download attachments from Office 365 emails and extract relevant information.

    Args:
        access_token (str): Access token obtained from the authentication process.
        start_date (date): Start date for filtering emails.
        start_time (time): Start time for filtering emails.

    Returns:
        tuple[dict, dict]: A tuple containing two dictionaries:
            - mail_dict: Dictionary containing extracted email information.
            - attachment_dict: Dictionary mapping email ID to attachment filenames.
    """
    path = os.path.join(os.getcwd(), "Attachments")
    os.makedirs(path, exist_ok=True)

    mail_dict: dict = {}
    attachment_dict: dict = {}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        # Fetch the shared mailbox emails using Microsoft Graph API
        url = "https://graph.microsoft.com/v1.0/users/sharedmailbox@domain.com/messages"
        params = {
            "$filter": f"receivedDateTime ge {start_date}T{start_time}Z",
            "$top": 100  # Adjust the page size as per your requirements
        }

        while True:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            messages = response.json()["value"]

            for message in messages:
                email_id = message["id"]

                # Fetch the email details
                url = f"https://graph.microsoft.com/v1.0/users/sharedmailbox@domain.com/messages/{email_id}"
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                email_data = response.json()

                attachments = []
                for attachment in email_data.get("attachments", []):
                    attachment_name = attachment["name"]
                    content_type = attachment["contentType"]

                    file_extension = os.path.splitext(attachment_name)[1].lower()
                    if file_extension == ".pdf":
                        attachments.append(attachment_name)

                        # Download the attachment
                        url = f"https://graph.microsoft.com/v1.0/users/sharedmailbox@domain.com/messages/{email_id}/attachments/{attachment['id']}/$value"
                        headers = {
                                        "Authorization": f"Bearer {access_token}",
                                        "Content-Type": "application/octet-stream"
                                    }
                        response = requests.get(url, headers=headers, stream=True)
                        response.raise_for_status()

                        attachment_path = os.path.join(path, attachment_name)
                        with open(attachment_path, "wb") as f:
                            f.write(response.content)

                # Extract relevant email information
                sender = email_data["sender"]["emailAddress"]["address"]
                date = email_data["receivedDateTime"]
                subject = email_data["subject"]
                body = email_data["body"]["content"]

                if email_data["body"]["contentType"] == "text/html":
                    body = extract_text_from_html(body)

                # Add email information to mail_dict only if it has PDF attachments
                if attachments:
                    mail_dict[email_id] = {
                        "sender": sender,
                        "date": date,
                        "subject": subject,
                        "body": body
                    }

                attachment_dict[email_id] = attachments

            # Check if there are more pages of emails
            next_link = response.json().get("@odata.nextLink")
            if not next_link:
                break

            url = next_link

        return mail_dict, attachment_dict

    except requests.exceptions.RequestException as e:
        print("An error occurred:", e)
        return mail_dict, attachment_dict

def filter_office365_attachments(mail_dict: dict, attachment_dict: dict) -> None:
    Subject_filter = {
        "[Not Virus Scanned] Ahli Bank -  Statement , Portfolio ": "Ahli",
        "[Not Virus Scanned] WOQOD - Fahes - Account Statement & Portfolio": "Fahes",
        "[Not Virus Scanned] Ahli Brokerage - Ardh Al Khaleej": "Waqod"
    }

    BodyTitle_filter = {
        'CIVIL FUND - QINVEST': 'Civil I',
        'GIRAFFA QIC': 'Giraffa',
        'OQIC': 'Oman',
        'MILITARY FUND - QINVEST / N': 'Military I',
        'QATAR INSURANCE COMPANY S.A.Q': 'QIC',
        'RASLAFFAN OPERATING COMPANY WLL': 'QEWC'
    }

    DocName_filter = {
        '170792': 'Pension C',
        '170793': 'Pension M'
    }

    path = os.path.join(os.getcwd(), "Attachments")

    for attachment in os.listdir(path):
        matching_keys = []
        for key, value in attachment_dict.items():
            if any(attachment in string for string in value):
                matching_keys.append(key)

        for key in matching_keys:
            def bodyTitleCheck() -> str | bool:
                for title in BodyTitle_filter:
                    if mail_dict[key]['body'].startswith(title):
                        return title
                return False

            def subjectCheck() -> str | bool:
                for subject in Subject_filter:
                    if mail_dict[key]['subject'].find(subject) != -1:
                        return subject
                return False

            if (sub := subjectCheck()):
                with open('PortfoliosPath.json') as file:
                    data = json.load(file)
                    destination_path = data[Subject_filter[sub]]
                    os.makedirs(destination_path, exist_ok=True)
                    shutil.copy(os.path.join(path, attachment), destination_path)
                    os.remove(os.path.join(path, attachment))

            elif (tc := bodyTitleCheck()):
                with open('PortfoliosPath.json') as file:
                    data = json.load(file)
                    destination_path = data[BodyTitle_filter[tc]]
                    os.makedirs(destination_path, exist_ok=True)
                    shutil.copy(os.path.join(path, attachment), destination_path)
                    os.remove(os.path.join(path, attachment))

            elif ((doc := os.path.splitext(attachment)[0]) in DocName_filter):
                with open('PortfoliosPath.json') as file:
                    data = json.load(file)
                    destination_path = data[DocName_filter[doc]]
                    os.makedirs(destination_path, exist_ok=True)
                    shutil.copy(os.path.join(path, attachment), destination_path)
                    os.remove(os.path.join(path, attachment))

            else:
                # Move to a separate folder for unclassified attachments
                unclassified_path = os.path.join(path, "Unclassified")
                os.makedirs(unclassified_path, exist_ok=True)
                shutil.move(os.path.join(path, attachment), unclassified_path)

    # Remove the "Attachments" folder if empty
    if not os.listdir(path):
        os.rmdir(path)


def main():
    client_id = getpass("Enter the Client ID (Application ID): ")
    client_secret = getpass("Enter the Client Secret (Application Secret): ")
    tenant_id = getpass("Enter the Tenant ID: ")

    start_date = datetime.strptime(input("Enter the start date (YYYY-MM-DD): "), "%Y-%m-%d").date()
    start_time = datetime.strptime(input("Enter the start time (HH:MM:SS): "), "%H:%M:%S").time()

    access_token = get_access_token(client_id, client_secret, tenant_id)

    mail_dict, attachment_dict = download_office365_attachments(access_token, start_date, start_time)

    filter_office365_attachments(mail_dict, attachment_dict)

    print("Attachments downloaded and filtered successfully!")


if __name__ == "__main__":
    main()
