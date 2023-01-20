import smtplib
import email.mime.application
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


fromaddr = 'jshenrique@assesi.com'
toaddr = "pinheiro012345@gmail.com"
msg = MIMEMultipart()


msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Contracheque Assesi"


body = MIMEText("""Segue em anexo o seu contracheque em PDF""")
msg.attach(body)


pdfname='cheque.pdf'
fp=open(pdfname,'rb')
anexo = email.mime.application.MIMEApplication(fp.read(),_subtype="pdf")
fp.close()
anexo.add_header('Content-Disposition','attachment',filename=pdfname)
msg.attach(anexo)


# iniciando o servidor
server = smtplib.SMTP('smtp.gmail.com', 587)    # setando conexçao com o host do gmail
server.starttls()   # iniciando a conexão
server.login(fromaddr, "mdyxmxrjqtxdaogz")  # logando com o email que vai enviar o email
text = msg.as_string()  # codificando as informações pro formato de leitura do gmail
server.sendmail(fromaddr, toaddr, text)  # enviando as informações
server.quit()   # fechando conexçao com o servidor

