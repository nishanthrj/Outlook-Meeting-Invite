import requests
import json
from datetime import datetime, timedelta
import csv
from decouple import config

CSV_FILE = "meeting.csv"

client_id: str = config("CLIENT_ID")
client_secret: str = config("CLIENT_SECRET")
tenant_id: str = config("TENANT_ID")
email_id = config("EMAIL")

access_token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
graph_api_endpoint = f"https://graph.microsoft.com/v1.0/users/{email_id}/events"


def get_access_token():
    token_payload = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }

    response = requests.post(access_token_url, data=token_payload)
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
        "attendees": [
            {
                "emailAddress": {
                    "address": data["To"],
                    "name": data["Name"],
                },
                "type": "Required",
            }
        ],
        "allowNewTimeProposals": False,
        "hideAttendees": True,
        "reminderMinutesBeforeStart": 30,
    }

    if data["Occurrence"] == "week":
        event_payload["recurrence"] = (
            {
                "pattern": {
                    "type": "daily",
                },
                "range": {
                    "type": "endDate",
                    "startDate": data["StartTime"],
                    "endDate": data["EndTime"],
                },
            },
        )

    response = requests.post(
        graph_api_endpoint, headers=headers, data=json.dumps(event_payload)
    )
    return response.json()


def main():
    access_token = get_access_token()

    if access_token:
        with open(CSV_FILE, "r") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                data = {
                    "To": row["To"],
                    "Name": row["Name"],
                    "Subject": row["Subject"],
                    "Body": row["Body"],
                    "Occurrence": row["Occurrence"],
                }

                try:
                    start_time = datetime.strptime(row["Date"], "%Y-%m-%dT%H:%M:%S")
                    duration = int(row["Duration"])
                    end_time = start_time + timedelta(
                        days=7 if row["Occurrence"] == "week" else 0, minutes=duration
                    )
                    data["StartTime"] = start_time.strftime("%Y-%m-%dT%H:%M:%S")
                    data["EndTime"] = end_time.strftime("%Y-%m-%dT%H:%M:%S")
                    response = create_event(access_token, data)
                    print(response)
                except KeyError as e:
                    print(f"Missing data for meeting '{data['Subject']}': {e}")

        print("Meeting invites sent successfully!")
    else:
        print("Failed to obtain access token.")


if __name__ == "__main__":
    main()
