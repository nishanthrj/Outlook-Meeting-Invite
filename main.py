import requests
import json
from datetime import datetime, timedelta
import csv
from decouple import config

CSV_FILE = "meeting.csv"

CLIENT_ID = config("CLIENT_ID")
CLIENT_SECRET = config("CLIENT_SECRET")
TENANT_ID = config("TENANT_ID")
EMAIL_ID = config("EMAIL")

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
EVENT_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ID}/events"
EMAIL_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ID}/sendMail"


def get_access_token():
    token_payload = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }

    response = requests.post(TOKEN_URL, data=token_payload)
    access_token = response.json().get("access_token")
    return access_token


def create_event(access_token, data):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    event_payload = {
        "subject": data["Subject"],
        "body": {
            "contentType": "HTML",
            "content": data["Body"],
        },
        "start": {
            "dateTime": data["StartTime"],
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": data["EndTime"],
            "timeZone": "UTC",
        },
        "recurrence": {
            "pattern": {"type": "daily", "interval": 1},
            "range": {
                "type": "endDate",
                "startDate": data["StartTime"][:10],
                "endDate": (
                    datetime.strptime(data["EndTime"][:10], "%Y-%m-%d")
                    + timedelta(days=7)
                ).strftime("%Y-%m-%d"),
            },
        },
        "attendees": [],
        "allowNewTimeProposals": False,
        "hideAttendees": True,
        "reminderMinutesBeforeStart": 30,
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
    }

    event_payload["attendees"] += [
        {
            "emailAddress": {
                "address": email,
                "name": name,
            },
            "type": "required",
        }
        for email, name in data["To"]
    ]

    event_payload["attendees"] += [
        {
            "emailAddress": {
                "address": email,
                "name": name,
            },
            "type": "optional",
        }
        for email, name in data["CC"]
    ]

    if data["Occurrence"] != "week":
        event_payload.pop("recurrence")

    response = requests.post(
        EVENT_ENDPOINT, headers=headers, data=json.dumps(event_payload)
    )

    return response.json()


def send_invites(access_token):
    with open(CSV_FILE, "r") as csvfile:
        reader = csv.DictReader(csvfile)
        groups = {}
        for row in reader:
            if row["Date"] in groups:
                groups[row["Date"]]["To"].append((row["To"], row["Name"]))
                groups[row["Date"]]["CC"].append((row["CCEmail"], row["CCName"]))

            else:
                groups[row["Date"]] = {
                    "To": [(row["To"], row["Name"])],
                    "CC": [(row["CCEmail"], row["CCName"])],
                    "Subject": row["Subject"],
                    "Body": row["Body"],
                    "Occurrence": row["Occurrence"],
                    "StartTime": row["Date"],
                    "EndTime": (
                        datetime.strptime(row["Date"], "%Y-%m-%dT%H:%M:%S")
                        + timedelta(minutes=int(row["Duration"]))
                    ).strftime("%Y-%m-%dT%H:%M:%S"),
                }

        for data in groups.values():
            response = create_event(access_token, data)
            error = response.get("error")
            if error:
                print(
                    f"Failed: Couldn't invite attendees for {data['Subject']}\n{error.get('message')}"
                )
            else:
                print(f"Success: Invited attendees for {data['Subject']}")


def send_feedback_email(access_token, to, cc, subject):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    email_payload = {
        "message": {
            "subject": f"Feedback Request for {subject}",
            "body": {
                "contentType": "Text",
                "content": f"We would appreciate your feedback on the {subject}. Thank you!",
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": email,
                        "name": name,
                    }
                }
                for name, email in to
            ],
            "ccRecipients": [
                {
                    "emailAddress": {
                        "address": email,
                        "name": name,
                    }
                }
                for name, email in cc
            ],
        },
        "saveToSentItems": "true",
    }

    response = requests.post(EMAIL_ENDPOINT, headers=headers, json=email_payload)
    return response


def ask_feedback(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    res = requests.get(EVENT_ENDPOINT + "?$select=subject,attendees", headers=headers)

    data = res.json()

    for event in data["value"]:
        subject = event["subject"]
        if "review" not in subject.lower():
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
        res = send_feedback_email(access_token, to, cc, subject)
        if res.status_code == 202:
            print(f"Success: Feedback request email sent successfully for {subject}")
        else:
            print(
                f"Failed: Couldn't send feedback request email for {subject}\nStatus code: {res.status_code}"
            )


def main():
    access_token = get_access_token()

    if access_token:
        send_invites(access_token)
        # ask_feedback(access_token)
    else:
        print("Failed to obtain access token.")


if __name__ == "__main__":
    main()
