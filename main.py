import win32com.client as win32
import datetime

date = datetime.datetime.today()
date = date.strftime("%m/%d/%y")
to_email = "WHO YOU WANT SEND EMAIL TO"

# Add your subject line and body of the message here
subject = "Assembly Language - YOUR NAME - I participated in class today on " + date
message = "I participated in class today.\n \n Thanks, \n YOUR NAME"

# Function that sends email


def sendEmail(to_email, subject, message):
    try:
        print("start try")
        outlook = win32.Dispatch('outlook.application')
        print("connected to outlook")
        mail = outlook.CreateItem(0)
        print("outlook created item")
        mail.To = to_email
        mail.Subject = subject
        mail.Body = message
        print("sending...")
        mail.Send()
        print("sent")
        return True
    except Exception as e:
        print("tried")
        print("Error: " + str(e))
        return False


sendEmail(to_email, subject, message)
