import win32com.client
from parse_mail import get_meeting_data
from settings import MEETING_RECIPIENT, OUTLOOK_APPOINTMENT_ITEM, OUTLOOK_MEETING, \
                     OUTLOOK_OPTIONAL_ATTENDEE, OUTLOOK_FORMAT, DURATION, REMIND

outlook = win32com.client.Dispatch('Outlook.Application')


def send_meeting(meeting_data):
    mtg = outlook.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
    mtg.MeetingStatus = OUTLOOK_MEETING
    mtg.Subject = meeting_data['node']
    mtg.Start = meeting_data['time_start'].strftime(OUTLOOK_FORMAT)
    mtg.Duration = DURATION
    mtg.ReminderMinutesBeforeStart = REMIND
    mtg.ResponseRequested = False
    mtg.Body = meeting_data['url']
    invite = mtg.Recipients.Add(MEETING_RECIPIENT)
    invite.Type = OUTLOOK_OPTIONAL_ATTENDEE
    print(f'Отправляю встречу {mtg.Subject} на {MEETING_RECIPIENT}')
    mtg.Send()


if __name__ == '__main__':
    meetings = get_meeting_data()
    for meeting in meetings:
        send_meeting(meeting)
