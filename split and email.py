import xxlimited
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from tkinter import filedialog
from tkinter import Tk
from email.mime.application import MIMEApplication

# Create file dialog and get file path
root = Tk()
root.withdraw()  # we don't want a full GUI, so keep the root window from appearing
file_path = filedialog.askopenfilename()  # show an "Open" dialog box and return the path to the selected file
root.update()  # To prevent tkinter's main loop from starting, this is necessary after using filedialog

# Load the excel file
df = pd.read_excel(file_path, sheet_name='your sheet name')


# Loop over unique respondents
for respondent in df.iloc[:, 10].unique():
    # change 10 to the actual colum of your data contains email address
    filtered_df = df[df.iloc[:, 10] == respondent]

    filtered_df.to_excel(f'{respondent}_divided.xlsx', index=False)
    # result = pd.concat([header,filtered_df])
    save_path = os.path.join(os.path.dirname(file_path), f'{respondent}_divided.xlsx')
    filtered_df.to_excel(save_path, index=False)
    

    
    # Email details
    mail_content = '''Hello,
    the data is as attached.
    '''
    # Setup the SMTP server and log in
    sender_address = 'Your email address'
    sender_pass = 'YOUR KEY'
    receiver_address = respondent
    smtp_server = smtplib.SMTP_SSL('smtp.xxx.com', xx) # update if necessary

    smtp_server.login(sender_address, sender_pass)
    
    # Setup MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = f'please enter your subject {respondent}'
    message.attach(MIMEText(mail_content, 'plain'))
    
    # Attach the .xlsx file
    attach_file_name = f'{respondent}_divided.xlsx'
    attach_file = open(attach_file_name, 'rb') # Open the file in bynary
    payload = MIMEBase('application', 'octate-stream')
    # Add payload header with pdf name
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload)
    
    payload.add_header('Content-Disposition', 'attachment', filename=attach_file_name)
    message.attach(payload)
    
    # Use smtp to send the email with the attachment
    text = message.as_string()
    smtp_server.sendmail(sender_address, receiver_address, text)
    smtp_server.quit()
    attach_file.close()
