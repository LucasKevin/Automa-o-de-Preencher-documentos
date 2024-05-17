from tkinter import *
import os
from docx import Document
from docx2pdf import convert
from tkinter import filedialog
from tkinter import messagebox

def preencher_documento():
    # Carregar o modelo do documento
    doc = Document('termo_headset.docx')
    
    # Dicionário com os campos do documento e suas respectivas entradas
    campos = {
        '<<NOME>>': nome_entry.get(),
        '<<CARGO>>': cargo_entry.get(),
        '<<RUA>>': rua_entry.get(),
        '<<BAIRRO>>': bairro_entry.get(),
        '<<CIDADE>>': cidade_entry.get(),
        '<<ESTADO>>': estado_entry.get(),
        '<<CEP>>': cep_entry.get(),
        '<<LOCAL_E_DATA>>': local_e_data_entry.get(),
        '<<RESPONSAVEL>>': responsavel_entry.get(),
    }

    # Iterar sobre o dicionário e substituir os campos no documento
    for campo, entrada in campos.items():
        if entrada:
            for p in doc.paragraphs:
                if campo in p.text:
                    p.text = p.text.replace(campo, entrada)

    # Salvar o doc preenchido como docx temporário
    temp_docx = 'headset_preenchido.docx'
    doc.save(temp_docx)

    # Converter o documento docx para pdf
    convert(temp_docx)  # Isso criará um arquivo 'documento_preenchido.pdf'

    # Salvar o documento preenchido
    doc.save('headset_preenchido.docx')

    # Limpar os campos do formulário
    limpar_campos()

def preencher_e_salvar():
    # Verificar se todos os campos foram preenchidos
    campos = [
        nome_entry.get(),
        cargo_entry.get(),
        rua_entry.get(),
        bairro_entry.get(),
        cidade_entry.get(),
        estado_entry.get(),
        cpf_entry.get(),
        local_e_data_entry.get(),
        responsavel_entry.get(),
    ]
    
    if all(campos):  # Se todos os campos foram preenchidos
        preencher_documento()

        # Renomear o arquivo PDF
        novo_nome = 'headset_preenchido.docx'

        # Mover o arquivo PDF para o diretório desejado
        # Certifique-se de especificar o diretório correto onde você deseja salvar o arquivo PDF
        novo_caminho = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=novo_nome, filetypes=[("PDF files", "*.pdf")])
        if novo_caminho:
            os.rename('headset_preenchido.pdf', novo_caminho)
    else:  # Se algum campo estiver vazio
        messagebox.showerror("Erro", "Por favor, preencha todos os campos antes de submeter.")


def limpar_campos():
    nome_entry.delete(0, 'end')
    cargo_entry.delete(0, 'end')
    rua_entry.delete(0, 'end')
    bairro_entry.delete(0, 'end')
    cidade_entry.delete(0, 'end')
    estado_entry.delete(0, 'end')
    cep_entry.delete(0, 'end')
    cpf_entry.delete(0, 'end')
    local_e_data_entry.delete(0, 'end')
    responsavel_entry.delete(0, 'end')

# Criar a janela principal
root = Tk()
root.title("Preencher Documento")

# Criar rótulos e campos para o formulário
Label(root, text="Nome:").grid(row=0, column=0, sticky=W)
nome_entry = Entry(root)
nome_entry.grid(row=0, column=1)

Label(root, text="Cargo:").grid(row=1, column=0, sticky=W)
cargo_entry = Entry(root)
cargo_entry.grid(row=1, column=1)

Label(root, text="Rua:").grid(row=2, column=0, sticky=W)
rua_entry = Entry(root)
rua_entry.grid(row=2, column=1)

Label(root, text="Bairro:").grid(row=3, column=0, sticky=W)
bairro_entry = Entry(root)
bairro_entry.grid(row=3, column=1)

Label(root, text="Cidade:").grid(row=4, column=0, sticky=W)
cidade_entry = Entry(root)
cidade_entry.grid(row=4, column=1)

Label(root, text="Estado:").grid(row=5, column=0, sticky=W)
estado_entry = Entry(root)
estado_entry.grid(row=5, column=1)

Label(root, text="CEP:").grid(row=6, column=0, sticky=W)
cep_entry = Entry(root)
cep_entry.grid(row=6, column=1)

Label(root, text="CPF:").grid(row=7, column=0, sticky=W)
cpf_entry = Entry(root)
cpf_entry.grid(row=7, column=1)

Label(root, text="Local e Data:").grid(row=8, column=0, sticky=W)
local_e_data_entry = Entry(root)
local_e_data_entry.grid(row=8, column=1)

Label(root, text="Responsável:").grid(row=9, column=0, sticky=W)
responsavel_entry = Entry(root)
responsavel_entry.grid(row=9, column=1)


# Botão de envio
Button(root, text="Preencher e Salvar", command=preencher_e_salvar).grid(row=10, column=0, columnspan=2)

# Iniciar o loop principal da interface gráfica
root.mainloop()
