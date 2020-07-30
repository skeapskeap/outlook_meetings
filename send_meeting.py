import win32com.client
import datetime as dt
from settings import MEETING_RECIPIENT, SRC_EMAIL

OUTLOOK_APPOINTMENT_ITEM    = 1
OUTLOOK_MEETING             = 1
OUTLOOK_ORGANIZER           = 0
OUTLOOK_OPTIONAL_ATTENDEE   = 2
OUTLOOK_FORMAT              = '%d/%m/%Y %H:%M'
ONE_HOUR                    = 60
FIFTEEN_MINUTES             = 15

outlook = win32com.client.Dispatch('Outlook.Application')


def send_meeting_request(subject, time, recipient, body):
    mtg = outlook.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
    mtg.MeetingStatus = OUTLOOK_MEETING
    mtg.Subject = subject
    mtg.Start = time.strftime(OUTLOOK_FORMAT)
    mtg.Duration = ONE_HOUR
    mtg.ReminderMinutesBeforeStart = FIFTEEN_MINUTES
    mtg.ResponseRequested = False
    mtg.Body = body
    invite = mtg.Recipients.Add(recipient)
    invite.Type = OUTLOOK_OPTIONAL_ATTENDEE
    mtg.Send()


if __name__ == "__main__":
    time = dt.datetime.now() + dt.timedelta(hours=3)
    test_recipient = MEETING_RECIPIENT
    test_sender = SRC_EMAIL

    send_meeting_request('Test Meeting', time, test_recipient, 'This is a test meeting.')
