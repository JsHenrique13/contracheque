import PySimpleGUI as sg
from pikepdf import Pdf # type: ignore  
import aspose.words as aw
import os
import smtplib
import email.mime.application
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


base_dir = os.getcwd()
try:
    os.mkdir('contracheques')
except:
    pass

def Criaarquivo(arquivo, destino):
    try:
        with open('config.txt', 'a+', encoding="UTF_8") as file:
            file.write(f"{arquivo}:{destino}\n")
    except Exception as e:
        sg.popup(e)


def Envia(arquivo, destino):

    fromaddr = 'jshenrique@assesi.com'
    toaddr = destino
    msg = MIMEMultipart()


    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "Contracheque Assesi"


    body = MIMEText("""Segue em anexo o seu Contracheque""")
    msg.attach(body)


    pdfname=arquivo
    fp=open(f"contracheques/{pdfname}",'rb')
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
    sg.PopupNoTitlebar(f"Contracheque enviado para {destino}")


def separador(file):
    # the target PDF document to split
    filename = os.path.basename(file)
    print(filename)
    # load the PDF file
    pdf = Pdf.open(filename)
    p = 1 
    
    for page in pdf.pages:

        novo = Pdf.new()
        novo.pages.append(page)
        
        name, ext = os.path.splitext(filename)
        nome_arq = f"{name}-{p}.pdf"
        novo.save(nome_arq)
        doc = aw.Document(nome_arq)    
        doc.save(f"{name}-{p}.txt")     
        text = open(f"{name}-{p}.txt", "r", encoding="UTF-8")
        nome = text.readlines()[5][:2]
        nome = nome.strip()
        if len(nome) == 1: nome = '0'+ nome
        text.close()   
        try:        
            os.rename(nome_arq, f'contracheques/{nome}.pdf')
            os.remove(f"{name}-{p}.txt")
        except FileExistsError:
            os.rename(nome_arq, f'contracheques/{nome}(1).pdf')
            os.remove(f"{name}-{p}.txt")
        p+=1
    

def Config():
    file = 'config.txt'
    base_dir = os.getcwd()
    # pasta = os.listdir(f'{base_dir}/contracheques')
    menu_header = [
    ['Opções',['Enviar', 'Menu Principal']],
    
    ]

    layout = [
            [sg.Menu(menu_header, text_color='black', pad=(5,5), font='Courier 10')],
            [sg.Text('Matrícula', size=(10,0), text_color='#d0cdd7', background_color='#7b0d1e', pad=(1,1)), sg.Text('Endereço de Email', size=(30,0), text_color='#d0cdd7', background_color='#7b0d1e',  pad=(1,1))],
            [sg.Button("Salvar Dados"), sg.Button("Voltar")]
    ]
    c=2
    open(file, 'r+')
    for linha in file.readlines():
        val = linha.split(sep=':')
        print(linha[0], linha[1])
        row_content = [
            [sg.Text(val[0] , size=(10,0), pad=(1,1), key=linha[0][:-1]), sg.Input(size=(30,0), pad=(1,1), key=linha[1])]
        ]
        layout.insert(c, row_content)
        c+=1


    return sg.Window("Configurar Destinatários", layout=layout, resizable=True , font='Helvetica 10', element_justification='c', text_justification='c', finalize=True)


def Main():

    menu_header = [
    ['Opções',['Configurar', 'Enviar']],
    
    ]

    layout = [
            
            [sg.Menu(menu_header, text_color='black', pad=(5,5), font='Courier 10')],
            [sg.VPush()],
            [sg.Text("Buscar arquivo de Contracheques...")],
            [sg.Input("", size=(30,1), do_not_clear=False, pad=(1,1)),sg.FileBrowse('Buscar Arquivo', file_types=(("PDF Files", "*.pdf"),), initial_folder=base_dir, pad=(0,0), font='Helvetica 12')],
            [sg.Text("")],
            [sg.Button("Separar PDF's", font='Helvetica 10'), sg.Button("Sair", font='Helvetica 10')],
    ]

    return sg.Window("Contracheque", layout=layout, resizable=True , font='Courier 14', element_justification='c', text_justification='c', finalize=True,)


telamain, telaconfig = Main(), None

while True:
    window, event, values = sg.read_all_windows()   

    if event in (sg.WIN_CLOSED, 'Sair'):
        break
    if window == telamain and event == 'Configurar':
        telamain.hide()
        telaconfig = Config()
    if window == telamain and event == "Separar PDF's":
        separador(values["Buscar Arquivo"])
    if window == telaconfig and event == "Salvar Dados":
        for item in values:
            Criaarquivo(item, values[item])
        
        telamain.UnHide()
        telaconfig.hide()
    if window == telamain and event == 'Enviar':
        base_dir = os.getcwd()
        pasta = os.listdir(f'{base_dir}/contracheques')
        for item in values:
            if values[item] != '':
                print(item, values[item])
        
    if window == telaconfig and event == "Voltar":
        telamain.UnHide()
        telaconfig.hide()
    




