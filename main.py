import json
import csv
import time
import requests
from decouple import config
from utils import (
    generate_group_id,
    get_recurrence_pattern,
    format_date,
    format_meeting_body,
)

# The main file with all meeting information.
CSV_FILE = "meetings.csv"

# API authentication details.
CLIENT_ID = config("CLIENT_ID")
CLIENT_SECRET = config("CLIENT_SECRET")
TENANT_ID = config("TENANT_ID")
OBJECT_ID = config("OBJECT_ID")

# API endpoints.
TOKEN_ENDPOINT = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
EVENT_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{OBJECT_ID}/events"
EMAIL_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{OBJECT_ID}/sendMail"

# The token required to make API calls.
ACCESS_TOKEN = None


def set_access_token():
    """
    Set the access token for Microsoft Graph API.
    """

    global ACCESS_TOKEN

    token_payload = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }

    response = requests.post(TOKEN_ENDPOINT, data=token_payload)
    ACCESS_TOKEN = response.json().get("access_token")


def create_event_payload(data: dict) -> dict:
    """
    Create payload for creating an Outlook event.

    Args:
        data (dict): Event data.

    Returns:
        dict: Payload for creating an event.
    """

    payload = {
        "subject": data["Subject"],
        "body": {
            "contentType": "HTML",
            "content": format_meeting_body(data["Body"], data["MeetingURL"]),
        },
        "start": {
            "dateTime": f"{data['StartDate']}T{data['StartTime']}",
            "timeZone": data["TimeZone"],
        },
        "end": {
            "dateTime": f"{data['StartDate']}T{data['EndTime']}",
            "timeZone": data["TimeZone"],
        },
        "recurrence": {
            "pattern": get_recurrence_pattern(data["Occurrence"], data["StartDate"]),
            "range": {
                "type": "endDate",
                "startDate": data["StartDate"],
                "endDate": data["EndDate"],
            },
        }
        if data["Occurrence"] != "once"
        else None,
        "attendees": [
            {"emailAddress": {"address": email, "name": name}, "type": "required"}
            for email, name in data["To"]
        ]
        + [
            {"emailAddress": {"address": email, "name": name}, "type": "optional"}
            for email, name in data["CC"]
        ],
        "allowNewTimeProposals": False,
        "hideAttendees": True,
        "reminderMinutesBeforeStart": 30,
        "isOnlineMeeting": True if data["Platform"] == "Teams" else False,
        "onlineMeetingProvider": "teamsForBusiness"
        if data["Platform"] == "Teams"
        else "unknown",
    }

    return payload


def create_event(data: dict) -> dict:
    """
    Create an event using Microsoft Graph API.

    Args:
        data (dict): Event data.

    Returns:
        dict: JSON response from the API.
    """

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    event_payload = create_event_payload(data)
    response = requests.post(
        EVENT_ENDPOINT, headers=headers, data=json.dumps(event_payload)
    )
    return response.json()


def send_event_invites() -> None:
    """
    Read data from a CSV file and send meeting invites.
    """

    groups = {}

    with open(CSV_FILE, "r") as csvfile:
        reader = csv.DictReader(csvfile)

        for row in reader:
            group_id = generate_group_id(
                row["StartDate"] + row["StartTime"] + row["EndDate"] + row["EndTime"]
            )
            if group_id in groups:
                groups[group_id]["To"].append((row["To"], row["Name"]))
                groups[group_id]["CC"].append((row["CCEmail"], row["CCName"]))
            else:
                groups[group_id] = {
                    "To": [(row["To"], row["Name"])],
                    "CC": [(row["CCEmail"], row["CCName"])],
                    "Subject": row["Subject"],
                    "Body": row["Body"],
                    "Occurrence": row["Occurrence"],
                    "StartDate": format_date(row["StartDate"]),
                    "EndDate": format_date(row["EndDate"]),
                    "StartTime": row["StartTime"],
                    "EndTime": row["EndTime"],
                    "TimeZone": row["TimeZone"],
                    "Platform": row["Platform"],
                    "MeetingURL": row["MeetingURL"],
                }

    for data in groups.values():
        response = create_event(data)
        error = response.get("error")
        if error:
            print(
                f"Failed: Couldn't invite attendees for {data['Subject']}\n{error.get('message')}"
            )
        else:
            print(f"Success: Invited attendees for {data['Subject']}")


def create_email_payload(to: list, cc: list, subject: str) -> dict:
    """
    Create payload for sending a feedback email.

    Args:
        to (list): List of to recipients.
        cc (list): List of cc recipients.
        subject (str): Email subject.

    Returns:
        dict: Payload for sending an email.
    """

    return {
        "message": {
            "subject": f"Feedback Request for {subject}",
            "body": {
                "contentType": "Text",
                "content": f"We would appreciate your feedback on the {subject}. Thank you!",
            },
            "toRecipients": [
                {"emailAddress": {"address": email, "name": name}} for name, email in to
            ],
            "ccRecipients": [
                {"emailAddress": {"address": email, "name": name}} for name, email in cc
            ],
        },
        "saveToSentItems": "true",
    }


def send_feedback_email(to: list, cc: list, subject: str) -> requests.Response:
    """
    Send a feedback email using Microsoft Graph API.

    Args:
        to (list): List of to recipients.
        cc (list): List of cc recipients.
        subject (str): Email subject.

    Returns:
        requests.Response: Response from the API.
    """

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }
    email_payload = create_email_payload(to, cc, subject)
    response = requests.post(EMAIL_ENDPOINT, headers=headers, json=email_payload)
    return response


def ask_feedback() -> None:
    """
    Retrieve events with 'feedback' in the subject and send feedback emails.
    """

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }

    res = requests.get(EVENT_ENDPOINT + "?$select=subject,attendees", headers=headers)
    data = res.json()

    for event in data["value"]:
        subject = event["subject"]
        if "feedback" not in subject.lower():
            continue

        to = []
        cc = []
        for attendee in event["attendees"]:
            name = attendee["emailAddress"]["name"]
            email = attendee["emailAddress"]["address"]

            if attendee["type"] == "required":
                to.append((name, email))
            elif attendee["type"] == "optional":
                cc.append((name, email))

        response = send_feedback_email(to, cc, subject)

        if response.status_code == 202:
            print(f"Success: Feedback request email sent successfully for {subject}")
        else:
            print(
                f"Failed: Couldn't send feedback request email for {subject}\nStatus code: {response.status_code}"
            )


def main():
    """
    Main function to execute the script.
    """

    set_access_token()

    if ACCESS_TOKEN:
        send_event_invites()
        # Wait for 10s before running the next part.
        time.sleep(10)
        ask_feedback()
    else:
        print("Failed to obtain access token.")


if __name__ == "__main__":
    main()
