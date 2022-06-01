import win32com.client
import time
import csv
from datetime import date, datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders["SUBFOLDER"]
messages = inbox.Items

# Modify code below to filter by time

# Filters by last 24 hours
# received_dt = datetime.now() - timedelta(days=1)
# received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
# messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

# Filter by specifying a date explicitly
messages = messages.Restrict("[ReceivedTime] >= 'April 01,'")

# Modify code below to filter by sender
# messages = messages.Restrict("[SenderEmailAddress] = 'postmaster@gov.mb.ca'")

# Modify code below to filter by subject
# messages = messages.Restrict("[Subject] = 'Password Protected File Blocked'")

# Uncomment code block for cross-checking the first email
# latest_message = messages.GetLast()
# index = latest_message.Body.index(">") + 1
# body = latest_message.Body[index:-3]
# print(body.split("\r\n"))

# Sorts the emails by the time it was received
messages.Sort("[ReceivedTime]", Descending=False)

# Uncomment this to get how many emails in total
# print(len(messages))

# Uncomment code block to print test output on the command line
# for message in messages:
#     index = message.Body.index(">") + 1
#     body = message.Body[index:-3]
#     if len(body.split("\r\n")) != 5:
#         sender = body.split("\r\n")[1]
#         recipient = body.split("\r\n")[2]
#         subject = "No Subject"
#     else:
#         sender = body.split("\r\n")[1]
#         recipient = body.split("\r\n")[2]
#         subject = body.split("\r\n")[3]
    
#     format_sender = "Sender: {:<15}"
#     format_recipient = "Recipient: {:<15}"
#     format_subject = "Subject: {:<15}"

#     print(format_sender.format(sender))
#     print(format_recipient.format(recipient))
#     print(format_subject.format(subject))
#     print("=========================================================")

# Exports a .csv file in the specified path using the csv python module
# Learn more about csv module here: https://docs.python.org/3/library/csv.html
with open('C:/PATH/TO/CSV_FILE.CSV', 'w') as CSV_FILE:
    writer = csv.writer(CSV_FILE)

    for message in messages:
        index = message.Body.index(">") + 1
        body = message.Body[index:-3]
        if len(body.split("\r\n")) != 5:
            sender = body.split("\r\n")[1]
            recipient = body.split("\r\n")[2]
            subject = "No Subject"
        else:
            sender = body.split("\r\n")[1]
            recipient = body.split("\r\n")[2]
            subject = body.split("\r\n")[3]

        writer.writerow([sender, recipient, subject])

# Prints how many emails were exported
print(len(messages))