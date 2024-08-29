import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


host = "smtp.gmail.com"
port = 587
login = "renapus00@gmail.com"
senha = "XXXX"


server = smtplib.SMTP(host, port)
server.ehlo()
server.starttls()
server.login(login, senha)


corpo = "<b>Bom dia</b> <p>Segue em anexo a planilha com as informações coletadas.</p><p>Atenciosamente,<br>Nome</p>"
email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login
email_msg['Subject'] = "ENVIO DE PLACAS"


email_msg.attach(MIMEText(corpo, 'html'))


anexo = "testecomIA.xlsx"
with open(anexo, 'rb') as file:
    parte = MIMEBase('application', 'octet-stream')
    parte.set_payload(file.read())


encoders.encode_base64(parte)


parte.add_header('Content-Disposition', f'attachment; filename="{anexo}"')
email_msg.attach(parte)

# Enviando o email
server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())

# Fechando a conexão com o servidor
server.quit()