"""
https://docs.microsoft.com/en-gb/office/vba/api/outlook.folder
https://docs.microsoft.com/en-us/archive/msdn-magazine/2013/march/powershell-managing-an-outlook-mailbox-with-powershell
https://devblogs.microsoft.com/premier-developer/outlook-email-automation-with-powershell/
https://community.spiceworks.com/how_to/150253-send-mail-from-powershell-using-outlook
"""

from datetime import datetime as dt, timedelta
import re
import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")  # Это магия чтобы обращаться к COM-объектам Outlook.Application
"""
    Эти штуки итерируются о_О
    for item in email_boxes:
        print(item)
"""
email_boxes = outlook.Folders  # Все ящики, добавленные в Outlook
my_box = email_boxes.Item(6)        # это мой личный ящик
my_inbox = my_box.Folders.Item(4)   # это папка Входящие в моём ящике
target_folder = my_inbox.Folders.Item('Notifications')  # Можно даже по имени папку указывать
all_emails = target_folder.Items


def parse_notification(notification: str) -> dict:
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


def proper_dt(com_dt):
    date_string = dt.strftime(com_dt, '%d.%m.%y %H:%M:%S')       # Преобразование datetime COM-объекта в строку
    standard_dt = dt.strptime(date_string, '%d.%m.%y %H:%M:%S')  # Преобразование строки в стандартный datetime
    return standard_dt


def last_hour_emails(emails) -> list:
    hour_ago = dt.now()-timedelta(hours=1)
    notifications = [item.Body for item in emails if proper_dt(item.SentOn) > hour_ago]
    return notifications


if __name__ == '__main__':
    for item in last_hour_emails(all_emails):
        print(parse_notification(item))
