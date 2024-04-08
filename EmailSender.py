import sys
import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

def send_email(file_path):

    # Create a new mail item
    mail = olApp.CreateItem(0)  # 0 represents the constant for a mail item

    # Set properties of the email
    mail.Subject = 'Test Email'
    mail.Body = 'This is a test email sent from Python'
    mail.To = 'moid.analyst@gmail.com'  # Replace with the actual recipient's email address

    attachment = file_path
    mail.Attachments.Add(attachment)

    # Send the email
    mail.Send()

    print('Email sent successfully!')

if __name__ == '__main__':
    file_path = sys.argv[1]
    send_email(file_path)







