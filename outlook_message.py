import win32com.client as win32


def create_message(send_From, send_To, action='display',
                   send_CC=None, send_BCC=None, subject=None,
                   body=None, html_body=None, attachment=None):
    OutlookApp = win32.Dispatch('outlook.application')
    mail = OutlookApp.CreateItem(0)
    
    mail.To = send_To
    if send_CC != None:
        mail.CC = send_CC
    
    if send_BCC != None:
        mail.BCC = send_BCC
    
    if subject != None:
        mail.Subject = subject
    
    if body != None:
        mail.Body = body
    
    if html_body != None:
        mail.HTMLBody = html_body
    
    if attachment != None:
        mail.Attachments.Add(attachment)
    
    try:
        mail.From = FROM
        if action.lower() == 'send':
            mail.Send()
        else:
            mail.Display()
        
    except AttributeError as error:
        mail.Display()
        print('Check if argument FROM is right:', error)
    

ATTACHMENT = "Attachment's path"
FROM = 'email@example.com'
TO = 'email@example.com'
SUBJECT = "Subject"
BODY = None
HTML_BODY = "<h1>This is example message</h1>"
    
create_message(FROM, TO, action='display', subject=SUBJECT, html_body=HTML_BODY)
