import smtplib
import email.mime.application
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


fromaddr = 'jshenrique@assesi.com'
toaddr = "pinheiro012345@gmail.com"
msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Contracheque Assesi"

"""
body = 

    <div style="@import url('https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,100;0,200;0,300;0,400;0,500;0,700;0,800;0,900;1,100;1,500;1,700;1,800&display=swap'); font-family: 'Poppins', sans-serif;     font-size: 1.2em;
    font-weight: 500; display: flex; justify-content: center; align-items: center; background-image: linear-gradient( 45deg, #FA4133, #94a406 ); text-align: justify; color: #fff; width: 250px; height: 250px; border-radius: 15px;" >

        <p >

            Segue em anexo o seu contracheque.

        </p>

    </div>
"""
body = email.mime.Text.MIMEText("""Segue em anexo o seu contracheque em PDF""")
msg.attach(body)

#   msg.attach(MIMEText(body, 'html'))
"""
"""
# adicionando o pdf
filename = 'JOSÉ HENRIQUE MARTINS PINHEIRO.pdf' # colocando o nome/path do arquivo em uma variavel

attachment = open('JOSÉ HENRIQUE MARTINS PINHEIRO.pdf', 'rb')  # abrindo o arquivo em modo leitura

part = MIMEBase('application', 'octet-stream')  # iniciando a função que vincula um item no email
part.set_payload(attachment.read())  # vinculando o arquivo desejado a este item
encoders.encode_base64(part)  # tranformando o item em base_64(formato reconhecido pela tercnologia de envio de emails)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)  # informando o diretorio do arquivo convertido

msg.attach(part)  # vinculando o anexo pronto na mensage a ser enviada

attachment.close()  # fechando a leitura do arquivo"""


pdfname='JOSÉ HENRIQUE MARTINS PINHEIRO.pdf'
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

