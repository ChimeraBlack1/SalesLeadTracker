import email
import imaplib
import base64
from bs4 import BeautifulSoup
from pwk import outlookUN
from pwk import outlookPW
import xlwt

username = outlookUN
password = outlookPW

mail = imaplib.IMAP4_SSL("outlook.office365.com")
mail.login(username, password)

mail.select("INBOX/Leads")

result, data = mail.uid('search', None, "ALL")
myData = data[0].split()

referralList = []

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
  # print(subject_)
  # print(text)

  

  # split text into List
  bodyTextList = text.split()
  
  # Clear out useless stuff
  try:
    please_ = bodyTextList.index("Please")
    bodyTextList.pop(please_)
    assign_ = bodyTextList.index("assign.")
    bodyTextList.pop(assign_)
  except:
    pass

  try:
    assign_ = bodyTextList.index("assign,")
    bodyTextList.pop(assign_)
  except:
    pass

  try:
    assign_ = bodyTextList.index("assign")
    bodyTextList.pop(assign_)
  except:
    pass

  # Company Details Merge
  try:
    companyIndex_ = bodyTextList.index("Company:")
  except:
    continue

  try:
    contactIndex_ = bodyTextList.index("Contact:")
  except:
    pass

  try:
    bodyTextList[companyIndex_: contactIndex_] = [' '.join(bodyTextList[companyIndex_+1: contactIndex_])]
  except:
    pass

  # Contact Details Merge
  try:
    contactIndex_ = bodyTextList.index("Contact:")
  except:
    pass

  try:
    phoneIndex_ = bodyTextList.index("Phone:")
  except:
    pass
  
  try:
    bodyTextList[contactIndex_: phoneIndex_] = [' '.join(bodyTextList[contactIndex_+1: phoneIndex_])]
  except:
    pass

  # Phone Details Merge
  try:
    phoneIndex_ = bodyTextList.index("Phone:")
  except:
    pass

  try:
    emailIndex_ = bodyTextList.index("Email:")
  except:
    pass
  
  try:
    bodyTextList[phoneIndex_: emailIndex_] = [' '.join(bodyTextList[phoneIndex_+1: emailIndex_])]
  except:
    pass

  # Email Details Merge
  try:
    emailIndex_ = bodyTextList.index("Email:")
  except:
    pass

  try:
    addyIndex_ = bodyTextList.index("Address:")
  except:
    pass  

  try:
    bodyTextList[emailIndex_: addyIndex_] = [' '.join(bodyTextList[emailIndex_+1: addyIndex_])]
  except:
    pass

  # Address Details Merge
  try:
    addyIndex_ = bodyTextList.index("Address:")
  except:
    pass  

  try:
    prodIndex_ = bodyTextList.index("Product")
  except:
    pass
  
  try:
    bodyTextList[addyIndex_: prodIndex_] = [' '.join(bodyTextList[addyIndex_+1: prodIndex_-1])]
  except:
    pass

  # Product Interest Details Merge
  try:
    prodIndex_ = bodyTextList.index("Product")
  except:
    pass

  try:
    descIndex_ = bodyTextList.index("Description:")
  except:
    pass

  try:
    bodyTextList[prodIndex_: descIndex_] = [' '.join(bodyTextList[prodIndex_+2: descIndex_])]
  except:
    pass

  # Description Details Merge
  try:
    descIndex_ = bodyTextList.index("Description:")
  except:
    pass

  try:
    jonIndex_ = bodyTextList.index("DeCiantis")
    jonIndex_ = jonIndex_ -1 
  except:
    pass

  try:
    bodyTextList[descIndex_: jonIndex_] = [' '.join(bodyTextList[descIndex_+1: jonIndex_])]
  except:
    pass

  # remove everything after description
  try:
    deCiantis_ = bodyTextList.index("DeCiantis")
  except:
    pass

  # convert list to object
  referral_ = {
    "Company": bodyTextList[0],
    "Contact": bodyTextList[1],
    "Phone": bodyTextList[2],
    "Email": bodyTextList[3],
    "Address": bodyTextList[4],
    "Product": bodyTextList[5],
    "desc": bodyTextList[6]
  }

  referralList.append(referral_)
  
  # clear everything after 'desc'
  # bodyTextList = bodyTextList[:deCiantis_-1]

for i in referralList:
  print("..............................")
  print(i)
  








  

  