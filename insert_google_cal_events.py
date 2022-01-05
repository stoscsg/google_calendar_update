"""
Insert Google Calendar Event
  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

Add credentials:
https://developers.google.com/calendar/api/guides/auth

"""
from __future__ import print_function

import datetime
import os.path
import csv
import rfc3339
import iso8601
from datetime import timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


def get_events():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        print("Getting the upcoming 10 events")
        events_result = (
            service.events()
            .list(calendarId="primary", timeMin=now, maxResults=10, singleEvents=True, orderBy="startTime")
            .execute()
        )
        events = events_result.get("items", [])

        if not events:
            print("No upcoming events found.")
            return

        # Prints the start and name of the next 10 events
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(start, event["summary"])

    except HttpError as error:
        print("An error occurred: %s" % error)


def insert_event(title, start, description):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build("calendar", "v3", credentials=creds)


    start_time = rfc3339.rfc3339(iso8601.parse_date(start))
    end_time = rfc3339.rfc3339(iso8601.parse_date(start) + timedelta(hours=3))


    event = {
        "summary": f"STOSC: {title}",
        "location": "650 Yio Chu Kang Rd, Singapore 787075",
        "description": description,
        "start": {"dateTime": start_time},
        "end": {"dateTime": end_time},
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup", "minutes": 1440},
            ],
        },
    }

    # calendarId='primary' for the main Calendar.
    # Get IDs from https://developers.google.com/calendar/api/v3/reference/calendarList/list
    event = service.events().insert(calendarId="xxx@group.calendar.google.com", body=event).execute()
    print(f"Event created: STOSC: {title} - [{start}]")


if __name__ == "__main__":
    # get_events()

    for row in csv.DictReader(open("insert_google_cal_event.csv")):
        insert_event(
            title=row["title"],
            start=f"{row['service_date_start']}T{row['start_time'].zfill(5)}+08:00",
            description=f"<b>{row['bible']}</b>\n\n{row['desc']}"
        )
