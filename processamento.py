import imaplib
import base64
import os
import email
import re

email_user = 'email_aqui'
email_pass = 'senha_aqui'

detach_dir = '.'
if 'olokinho' not in os.listdir(detach_dir):
    os.mkdir('olokinho')
mail = imaplib.IMAP4_SSL("Outlook.office365.com", 993)

mail.login(email_user, email_pass)

mail.select('Inbox')

typ, data = mail.search(None, 'ALL')
mail_ids = data[0]
id_list = mail_ids.split()
for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            explod_file = fileName.split('.')
            not_special_char = re.sub(u'[^a-zA-Z0-9áéíóúÁÉÍÓÚâêîôÂÊÎÔãõÃÕçÇ ]', '', explod_file[0])
            if len(explod_file) >= 2:
                print(explod_file)
                filePath = os.path.join(detach_dir, 'olokinho', not_special_char + '.' + explod_file[1])
                if not os.path.isfile(filePath):
                    fp = open(filePath, "wb")
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
                print('Arquivo: "{file}" baixado do email "{subject}"'.format(file=fileName, subject=subject))
