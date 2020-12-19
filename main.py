import smtplib
import pandas as pd


data = pd.read_excel('email_data.xlsx')
data_dict = data.to_dict('list')
emails = data_dict['emails']
messages = data_dict['messages']



combo = zip(emails, messages)
sender_email = "microintelligence786@gmail.com"
password = str("khanamir@786")
first = True


for emails, messages in combo:
    rec_email = emails
    message = messages
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, password)
    print("login sucessfuly")
    server.sendmail(sender_email, rec_email, message)
    print("Email has been sent to ", rec_email)
