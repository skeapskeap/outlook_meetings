import argparse
import sys
import win32com.client
from parse_mail import get_meeting_data
from settings import MEETING_RECIPIENT, OUTLOOK_APPOINTMENT_ITEM
from settings import OUTLOOK_MEETING, OUTLOOK_OPTIONAL_ATTENDEE, OUTLOOK_FORMAT
from settings import DURATION, REMIND, USER, FOLDER


arg_parser = argparse.ArgumentParser(
    description='send_meeting.exe -u <user@domain.ru> -f <notification folder>')
arg_parser.add_argument('-u', '--user', nargs='?', default=USER, type=str)
arg_parser.add_argument('-f', '--folder', nargs='?', default=FOLDER, type=str)
args = arg_parser.parse_args(sys.argv[1:])
user = args.user
folder = args.folder

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
    meetings = get_meeting_data(user=user, folder=folder)
    for meeting in meetings:
        if meeting:
            send_meeting(meeting)
        else:
            input('Press any key...')
