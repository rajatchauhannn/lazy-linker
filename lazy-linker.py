import imaplib
import email
import re
word = ["Click here to join the meeting"]

mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login('rajat12345chauhan@gmail.com', '240206rajat')

mail.list()
mail.select("Inbox", readonly=True)


for i in range(1,5):

    result, data = mail.uid('search', None, "ALL")

    latest_email_uid = data[0].split()[-i]

    result, data = mail.uid('fetch', latest_email_uid, '(RFC822)')

    raw_email = data[0][1]

    start_index = str(raw_email).find('When:')
    stop_index = str(raw_email).find('Where')

    print(str(raw_email)[start_index:stop_index])

    msg = email.message_from_string(str(raw_email))

    date_tuple = email.utils.parsedate_tz(msg['Date'])


    print(re.search("(?P<url>https?://[^\s]+)", str(raw_email)).group("url"))