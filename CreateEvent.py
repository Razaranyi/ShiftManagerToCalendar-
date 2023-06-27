from pprint import pprint
import app
from Google import Create_Service, convert_to_RFC_datetime

service = Create_Service(app.CLIENT_SECRET_FILE, app.API_NAME, app.API_VERSION, app.SCOPES)
print(dir(service))
calendar_ID = 'primary'
HOUR_ADJUSTMENT = -3


def insertEvents(latestSchedule):
    i = 1
    while i < len(latestSchedule):
        event_request_body = {
            'summary': latestSchedule[i].summary
            ,
            'start': {
                'dateTime': convert_to_RFC_datetime(latestSchedule[i].year,
                                                    latestSchedule[i].startMonth,
                                                    latestSchedule[i].startDay,
                                                    latestSchedule[i].startHour + HOUR_ADJUSTMENT,
                                                    latestSchedule[i].startMinutes),
                'timeZone': 'Asia/Jerusalem'
            },
            'end': {
                'dateTime': convert_to_RFC_datetime(latestSchedule[i].year, latestSchedule[i].endMonth,
                                                    latestSchedule[i].endDay,
                                                    latestSchedule[i].endHour + HOUR_ADJUSTMENT,
                                                    latestSchedule[i].endMinutes),
                'timeZone': 'Asia/Jerusalem'
            },
            'colorID': 5,
            'status': 'confirmed',
            'visibility': 'public',

        }
        response = service.events().insert(
            calendarId=calendar_ID,
            body=event_request_body
        ).execute()
        pprint(response)

        i = i + 1
