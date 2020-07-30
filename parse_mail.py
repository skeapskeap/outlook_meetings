"""
https://docs.microsoft.com/en-gb/office/vba/api/outlook.folder
https://docs.microsoft.com/en-us/archive/msdn-magazine/2013/march/powershell-managing-an-outlook-mailbox-with-powershell
"""

import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")

email_boxes = outlook.Folders  # Все ящики, добавленные в Outlook
"""
    эти штуки итерируются
    for item in email_boxes:
        print(item)
"""

my_box = email_boxes.Item(6)        # это моя личный ящик
my_inbox = my_box.Folders.Item(4)   # это папка Входящие в моём ящике
target_folder = my_inbox.Folders.Item('Notifications')  # Можно даже по имени папку указывать

emails = target_folder.Items
print(emails(1).Subject)
print(emails(1).SentOn)
print(emails(1).Body)
