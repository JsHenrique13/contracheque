import PySimpleGUI as sg
from pikepdf import Pdf # type: ignore  
import aspose.words as aw
import os


base_dir = os.getcwd()
try:
    os.mkdir('contracheques')
except:
    pass

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
        doc = aw.Document(nome_arq)     # type: ignore  
        doc.save(f"{name}-{p}.txt")     # type: ignore  
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
    print("OK")

def Config():
    base_dir = os.getcwd()
    tam = os.listdir(f'{base_dir}/contracheques')

    layout = [
            
            [sg.Text('Matrícula', size=(10,0), text_color='#d0cdd7', background_color='#7b0d1e', pad=(1,1)), sg.Text('Endereço de Email', size=(30,0), text_color='#d0cdd7', background_color='#7b0d1e',  pad=(1,1))],
            [sg.Button("Salvar Dados"), sg.Button("Voltar")]
    ]
    c=1
    for item in tam:
        val = item.split(sep='.')
        row_content = [
            [sg.Text(val[0] , size=(10,0), pad=(1,1)), sg.Input(size=(30,0), pad=(1,1))]
        ]
        layout.insert(c, row_content)
        c+=1


    return sg.Window("Configurar Destinatários", layout=layout, resizable=True , font='Helvetica 12', element_justification='c', text_justification='c', finalize=True)


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
    window, event, values = sg.read_all_windows()   # type: ignore

    if event in (sg.WIN_CLOSED, 'Sair'):
        break
    if window == telamain and event == 'Configurar':
        telamain.disappear()
        telaconfig = Config()
    if window == telamain and event == "Separar PDF's":
        separador(values["Buscar Arquivo"])

    if window == telaconfig and event == "Voltar":
        telamain.reappear()
        telaconfig.disappear()
    

print(event,values)




