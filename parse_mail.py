"""
https://docs.microsoft.com/en-gb/office/vba/api/outlook.folder
https://docs.microsoft.com/en-us/archive/msdn-magazine/2013/march/powershell-managing-an-outlook-mailbox-with-powershell
https://devblogs.microsoft.com/premier-developer/outlook-email-automation-with-powershell/
https://community.spiceworks.com/how_to/150253-send-mail-from-powershell-using-outlook
"""

from datetime import datetime as dt, timedelta
from settings import FOLDER, TIME_SPAN, USER
import re
import win32com.client, win32timezone


def get_notifications(user=USER, folder=FOLDER):
    # Это магия чтобы обращаться к COM-объектам Outlook.Application
    outlook = win32com.client.Dispatch(
        'Outlook.Application').GetNamespace("MAPI")
    """
        Эти штуки итерируются о_О
        for item in email_boxes:
            print(item)
    """
    # Все ящики, добавленные в Outlook
    email_boxes = outlook.Folders
    # это мой личный ящик
    my_box = email_boxes.Item(user)
    # это папка Входящие в моём ящике
    my_inbox = my_box.Folders.Item('Входящие')
    # Можно даже по имени папку указывать
    target_folder = my_inbox.Folders.Item(folder)
    return target_folder.Items


def get_meeting_data(**kwargs) -> list:
    notifications = get_notifications(**kwargs)
    meetings = [parse_notification(item)
                for item in recent_emails(notifications)]
    return meetings


def parse_notification(notification: str) -> dict:
    try:
        rows = notification.split('\n')
        node_address = re.split(r'[«»]', rows[0])[1]
        url = re.split(r'[<>]', rows[0])[1]
        times = re.findall(r'\d{2}:\d{2}', rows[3])
        dates = re.findall(r'\d{2}.\d{2}.\d{2}', rows[3])

        time_start = dates[0] + ' ' + times[0]
        time_start = dt.strptime(time_start, '%d.%m.%y %H:%M')

        meeting_data = {'node':         node_address,
                        'time_start':   time_start,
                        'url':          url}
        return meeting_data
    except IndexError:
        print('Письмо не парсится')
        return False


def proper_dt(com_dt):
    date_string = dt.strftime(com_dt, '%d.%m.%y %H:%M:%S')       # Преобразование datetime COM-объекта в строку
    standard_dt = dt.strptime(date_string, '%d.%m.%y %H:%M:%S')  # Преобразование строки в стандартный datetime
    return standard_dt


def recent_emails(emails) -> list:
    time_ago = dt.now()-timedelta(hours=TIME_SPAN)
    notifications = [item.Body for item in emails if proper_dt(item.SentOn) > time_ago]
    return notifications


if __name__ == '__main__':
    print(get_meeting_data())
