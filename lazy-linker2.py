import sys, win32com.client, datetime
# Connect with MS Outlook - must be open.
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace
("MAPI")
# connect to Sent Items
s = outlook.GetDefaultFolder(5).Items   # "5" refers to the sent item of a 
folder
#s.Sort("s", true)
# Get yesterdays date for the purpose of getting emails from this date
d = (datetime.date.today() - datetime.timedelta (days=1)).strftime("%d-%m-%
y")
# get the email/s
msg = s.GetLast()
# Loop through emails
while msg:
    # Get email date 
    date = msg.SentOn.strftime("%d-%m-%y")
    # Get Subject Line of email
    sjl = msg.Subject
    # Set the critera for whats wanted                       
    if d == date and msg.Subject.startswith("xx") or msg.Subject.startswith
    ("yy"):
    print("Subject:     " + sjl + "     Date : ", date) 
    msg = s.GetPrevious() 