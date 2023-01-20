import os
from pikepdf import Pdf
import aspose.words as aw


# the target PDF document to split
filename = "cheque.pdf"

# load the PDF file
pdf = Pdf.open(filename)
p = 1 
 
for n, page in enumerate(pdf.pages):

    novo = Pdf.new()
    novo.pages.append(page)
    
    name, ext = os.path.splitext(filename)
    nome_arq = f"{name}-{p}.pdf"
    novo.save(nome_arq)
    doc = aw.Document(nome_arq)
    doc.save(f"{name}-{p}.txt")
    text = open(f"{name}-{p}.txt", "r", encoding="UTF-8")
    nome = text.readlines()[5][:35]
    nomef = ''
    for n in nome: 
        if not n.isnumeric() :
            nomef += n
    nomef = nomef.strip()        
    try:        
        os.rename(nome_arq, f'{nomef}.pdf')
    except FileExistsError:
        os.rename(nome_arq, f'{nomef}{p}.pdf')
    p+=1


