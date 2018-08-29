"""
Read From Outlook Inbox

Author: eli


Reference link:
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem?view=outlook-pia

"""

import win32com.client


f_txt = open('inbox.txt', 'w+', encoding='utf-8')
f_html = open('inbox.html', 'w+', encoding='utf-8')

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')


# deleted_items = outlook.GetDefaultFolder(3)
# outbox = outlook.GetDefaultFolder(4)
# send_items = outlook.GetDefaultFolder(5)
inbox = outlook.GetDefaultFolder(6)

messages = list(inbox.Items)

for message in messages:
	print(message.subject, message.CreationTime)
	f_txt.write(f'Subject: {message.subject}\n')
	f_txt.write(f'Creation Time: {message.CreationTime}\n')
	f_txt.write(f'Received Time: {message.ReceivedTime}\n')
	f_txt.write(f'Sender: {message.SenderName}\n')
	f_txt.write('Recipients: {} \n'.format('; '.join([recipient.name for recipient in message.Recipients])))
	f_txt.write(f'Body: {message.body}\n')
	f_txt.write('\n-------------------------------------------------------------------------------\n')

f_txt.close()

for message in messages[:10]:
	print(message.subject, message.CreationTime)
	f_html.write(f'Subject: {message.subject}<br>')
	f_html.write(f'Creation Time: {message.CreationTime}<br>')
	f_html.write(f'Received Time: {message.ReceivedTime}<br>')
	f_html.write(f'Sender: {message.SenderName}<br>')
	f_html.write('Recipients: {} <br>'.format('; '.join([recipient.name for recipient in message.Recipients])))
	f_html.write(f'Body: {message.htmlbody}<br>')
	f_html.write('<hr>\n<hr>\n')

f_html.close()
