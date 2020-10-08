import imaplib
import base64
import os
import email
import re
from datetime import datetime
import time

# credenciais
email_user = 'email_aqui'
email_pass = 'senha_aqui'

# Cira um diretório caso não haja
detach_dir = '.'
if 'nome_pasta' not in os.listdir(detach_dir):
    os.mkdir('nome_pasta')

# conexão com servidor IMAP4
mail = imaplib.IMAP4_SSL("Outlook.office365.com", 993)
# realiza o login
mail.login(email_user, email_pass)
# acessa emails na caixa de entrada
mail.select('Inbox')

#seleciona emails de um único remetente ou de todos 'ALL'
typ, data = mail.search(None, 'FROM', '"email_remetente"')
# recupera os ids dos emails
mail_ids = data[0]
# cria uma lista de ids
id_list = mail_ids.split()
# percorre os emails
for num in data[0].split():
    # espera 1 segundo para executar
    time.sleep(1)
    # recupera a data e hora atual
    now = datetime.now()
    # recupera a hora em que o arquivo foi processado
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    # remove caracteres que possivelmente causarão erro
    dt_string = dt_string.replace('/', '_').replace(':', '_').replace(' ', '_')
    # acessa a caixa de um id específico de email
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
    # decodifica o o email para ficar mais legível pela aplicação
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    # percorre os emails decodificados
    for part in email_message.walk():
        # verifica se tem anexo, se tiver continua
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        # recupera o nome do arquivo
        fileName = part.get_filename()
        # caso haja nome, processa
        if bool(fileName):
            # separa a extensão do arquivo
            explod_file = fileName.split('.')
            # remove caracteres especiais para evitar erros
            not_special_char = re.sub(u'[^a-zA-Z0-9áéíóúÁÉÍÓÚâêîôÂÊÎÔãõÃÕçÇ ]', '', explod_file[0])
            # verifica se existe extensão (alguns arquivos não possuem, esses retornam erro e impedem que os demais sejam processados)
            if len(explod_file) >= 2:
                # caso haja arquivos com nomes iguais esta linha corrige adicionando o datetime
                not_special_char = not_special_char + dt_string
                # junta o nome do arquivo com sua respectiva extensão
                filePath = os.path.join(detach_dir, 'olokinho', not_special_char + '.' + explod_file[1])
                #caso nao exista o arquivo no caminho indicado, segue
                if not os.path.isfile(filePath):
                    # baixa o arquivo
                    fp = open(filePath, "wb")
                    # decodifica para a extensão do arquivo
                    fp.write(part.get_payload(decode=True))
                    # fecha a decodificação e arquivo
                    fp.close()
                # exibe o nome do arquivo em output e o titulo do email
                # subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
                # print('Arquivo: "{file}" baixado do email com titulo "{subject}"'.format(file=fileName, subject=subject))
