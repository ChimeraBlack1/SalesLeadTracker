import email
import imaplib
import base64
from bs4 import BeautifulSoup
from passwordKeeper import outlookUN
from passwordKeeper import outlookPW

username = outlookUN
password = outlookPW

mail = imaplib.IMAP4_SSL("outlook.office365.com")
mail.login(username, password)

mail.select("INBOX/Leads")

result, data = mail.uid('search', None, "ALL")
myData = data[0].split()

for item in myData:
  result2, email_data = mail.uid('fetch', item, '(RFC822)')
  raw_email = email_data[0][1].decode("utf-8")
  email_message = email.message_from_string(raw_email)
  to_ = email_message['To']
  from_ = email_message['From']
  subject_ = email_message['Subject']
  counter = 1
  for part in email_message.walk():
    if part.get_content_maintype() == "multipart":
      continue
    filename = part.get_filename()
    if not filename:
      ext = '.html'
      filename = 'msg-part-%08d%s' %(counter, ext)
      counter += 1

  content_type = part.get_content_type()
  html_ = part.get_payload()
  html_ = base64.b64decode(html_)
  soup = BeautifulSoup(html_, "html.parser")
  text = soup.get_text()
  print(subject_)
  print(text)




