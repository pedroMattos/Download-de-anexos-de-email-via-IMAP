# DOCUMENTACAO DO PYTHON IMAP4
# https://docs.python.org/2/library/imaplib.html#imaplib.IMAP4.select

import imaplib
import base64
import os
import email
import re
from datetime import datetime
import time
import threading

# credenciais
email_user = 'email'
email_pass = 'senha'

# Cira um diretorio caso nao haja
# windows
detach_dir = './backend/uploads'
# linux
# detach_dir = 'uploads'
# if './backend/nome_pasta' not in os.listdir(detach_dir):
#     os.mkdir('./backend/nome_pasta')

# conexao com servidor IMAP4
mail = imaplib.IMAP4_SSL("Outlook.office365.com", 993)
# realiza o login
mail.login(email_user, email_pass)
# acessa emails na caixa de entrada
# print(mail.list())
mail.select('"Nome da pasta"')
# mail.select('ALL')
remetente = ["email@email.com"]

def downloadFromEmail(remetente, dataMesPassado, mesAtual):
    print(remetente, ' Desde ', dataMesPassado, ' Até ', mesAtual)
    # seleciona emails de um único remetente ou de todos 'ALL'
    typ, data = mail.search(None, f'(FROM "{remetente}") (SINCE "{dataMesPassado}") (BEFORE "{mesAtual}")')
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
        # remove caracteres que possivelmente causarao erro
        dt_string = dt_string.replace('/', '_').replace(':', '_').replace(' ', '_')
        # acessa a caixa de um id especifico de email
        typ, data = mail.fetch(num, '(RFC822)' )
        raw_email = data[0][1]
        # decodifica o o email para ficar mais legivel pela aplicacao
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
                time.sleep(10)
                # separa a extensao do arquivo
                explod_file = fileName.split('.')
                # remove caracteres especiais para evitar erros
                not_special_char = re.sub(u'[^a-zA-Z0-9áéíóúÁÉÍÓÚâêîôÂÊÎÔãõÃÕçÇ ]', '', explod_file[0])
                # verifica se existe extensao (alguns arquivos nao possuem, esses retornam erro e impedem que os demais sejam processados)
                if len(explod_file) >= 2:
                    # caso haja arquivos com nomes iguais esta linha corrige adicionando o datetime
                    not_special_char = not_special_char
                    # junta o nome do arquivo com sua respectiva extensao
                    # windows
                    filePath = detach_dir + '/' + not_special_char + '.' + explod_file[1] + '.ZIP'
                    # linux
                    # filePath = detach_dir + '/' + not_special_char + '.' + explod_file[1] + '.z'
                    #caso nao exista o arquivo no caminho indicado, segue
                    if not os.path.isfile(filePath):
                        # baixa o arquivo
                        fp = open(filePath, "wb")
                        # decodifica para a extensao do arquivo
                        fp.write(part.get_payload(decode=True))
                        print('Arquivo baixado com sucesso!', filePath)
                        # fecha a decodificacao e arquivo
                        fp.close()
                    else:
                        print('Arquivo já existente', filePath)


now = datetime.now()
last_month = int(str(datetime.now())[5:7]) - 1
dt = now.strftime("%d-Dec-%Y")
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
hour = int(dt_string[11:13])
month = ['Jan', 'Feb', 'Mar', 'Apr', "May", "Jun", 'Jul','Aug', 'Sep', 'Oct', 'Nov', 'Dec']
mes = int(now.strftime("%m"))
dt = now.strftime(f"01-{month[last_month - 1]}-%Y")
ldt = now.strftime(f"28-{month[last_month]}-%Y")

def setInterval(func,time):
    e = threading.Event()
    while not e.wait(time):
        func()

def foo():
    if hour == 4 or hour == 12 or hour == 16 or hour == 22:
        print('Hora da Verificação no Email!')
        for i in remetente:
            downloadFromEmail(i, dt, ldt)
    else:
        print('A verificação no servidor acontecerá sempre às 04:00, 12:00, 14:00 e 22:00')

setInterval(foo, 10)
