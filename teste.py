import os
from pikepdf import Pdf # type: ignore  
import aspose.words as aw


try:
    os.mkdir('contracheques')
except:
    pass

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


