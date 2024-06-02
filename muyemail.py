import smtplib
import os
os.system('color a')

def smtp_live(email, activity, hour_count):
    '''
    Sends confirmation message to email specified confirming that the hour_count for the activity has been recorded in the Google Sheets
    '''

    username = "ialakash@live.com" #put your email here MUST BE AN OUTLOOK or LIVE email
    # password = PASSWORD_STRING
    smtp_server = "smtp-mail.outlook.com:587"
    # email_from = EMAIL_STRING
    email_to = email
    SUBJECT = 'Hours'
    TEXT = f"Thank you for submitting your hours for your activity {activity}. I have added {hour_count} hours to your record."
    message = f"Subject: {SUBJECT}\n\n{TEXT}"
 
    server = smtplib.SMTP(smtp_server)
    server.starttls()
    server.login(username, password)
    server.sendmail(email_from, email_to, message)
    server.quit()

if __name__ == '__main__':
    pass
