# Outlook Email Exporter

> Loops through all emails from an outlook mailbox and grabs their information such as sender, recipient, body, date, etc.

## Code

### 1. Import required python modules
```python
import win32com.client
import time
import csv
from datetime import date, datetime, timedelta
```

### 2. Create variables

```python
# Grabs the Outlook application using the win32 module
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Grabs the Inbox (6) folder or a subfolder (SUBFOLDER) within Inbox
inbox = outlook.GetDefaultFolder(6).Folders["SUBFOLDER"]

# Grabs the emails in the Inbox or specified SUBFOLDER
messages = inbox.Items

# Sorts the emails by the time it was received (oldest to newest)
messages.Sort("[ReceivedTime]", Descending=False)

```

### 3. Export .csv file

```python
with open('C:/PATH/TO/CSV_FILE.CSV', 'w') as CSV_FILE:
    writer = csv.writer(CSV_FILE)

    # Loops through each message in messages
    for message in messages:

        # Grabs the index of the ">" character's and adds 1 
        index = message.Body.index(">") + 1

        # Uses the index above to grab the relevant part of the message body
        body = message.Body[index:-3]

        # Splits the message body into sender, recipient, and subject line
        # Also checks if there is a sender and a subject
        if len(body.split("\r\n")) != 5:
            sender = body.split("\r\n")[1]
            recipient = body.split("\r\n")[2]
            subject = "No Subject"
        else:
            sender = body.split("\r\n")[1]
            recipient = body.split("\r\n")[2]
            subject = body.split("\r\n")[3]

        # Creates a row with the sender, recipient, and subject for each email found in the body
        writer.writerow([sender, recipient, subject])
```