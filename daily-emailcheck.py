import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta
import re

regex_name = r"(GC\-01 T)[^\n]*"
regex_ID = r"Session\sID:\s(\d{3}\-\d{3}\-\d{3})"
regex_closes = r"Bluebeam\sSession\sCloses\s(\d+/\d+/\d+)"


outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

account = ""
for account in mapi.Accounts:
	print(account.DeliveryStore.DisplayName)
	account = account.DeliveryStore.DisplayName

inbox = mapi.Folders(1).Folders("Inbox")


for idx, folder in enumerate(mapi.Folders(1).Folders):
    print(idx+1, folder)


messages = inbox.Items

today = datetime.now()
today = today.strftime('%m/%d/%Y')

received_dt = datetime.now() - timedelta(days=3)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderName] = 'shopdrawingsender'")

#place to save
outputDir = r"C:\Users\..."

output = []

try:
	message_list = list(messages)
	print(f"number of messages from sender {len(message_list)}")
	for message in message_list:
		print(f"\nsubject: {message.Subject}")
		try:
			email = message.Body
			name = re.search(regex_name, email)
			if name is not None:
				name = name.group(0).strip()
				print(name)
			else:
				name = message.Subject
				print(name)

			ID = re.search(regex_ID, email)
			if ID is not None:
				ID = ID.group(1)
				print(ID)
			else:
				continue

			closes = re.search(regex_closes, email)
			if closes is not None:
				closes = closes.group(1)
				print(closes)
			else:
				closes = "N/A"
				print("N/A")
			
			received_date = message.ReceivedTime
			received_date = received_date.strftime('%m/%d/%Y')
			print(received_date)
			message.Unread = False
			
			output.append(f"{name},{ID},{closes},{received_date}")
	    
		except Exception as e:
			print("error when saving the attachment:" + str(e))
except Exception as e:
		print("error when processing emails messages:" + str(e))

for x in output:
	print(x)

with open(outputDir, "w") as out:
	out.write("\n".join(output))
