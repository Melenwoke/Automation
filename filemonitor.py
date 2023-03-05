import os
import time
import win32com.client


with open('test1.txt', 'w') as f:
    f.write('This is a test file.')

# Define the path to the file to be monitored
file_path = u'test1.txt'  

# Define the email details
email_to = 'Michael.elenwoke@scrdairy.com'
email_subject = 'File updated'
email_body = 'The file has been updated.'

# Get the last modification time of the file
last_modified = os.path.getmtime(file_path)

while True:
    # Sleep for 6 seconds
    time.sleep(6)

    # Get the current modification time of the file
    current_modified = os.path.getmtime(file_path)

    # Check if the file has been modified
    if current_modified != last_modified:
        print('File has been modified.')
        
        # Create the Outlook email object
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        # Set the email properties
        mail.To = email_to
        mail.Subject = email_subject
        mail.Body = email_body

        # Attach the file to the email
        attachment = mail.Attachments.Add(file_path)

        # Send the email
        mail.Send()
        print('Email sent.')

        # Update the last modification time
        last_modified = current_modified
