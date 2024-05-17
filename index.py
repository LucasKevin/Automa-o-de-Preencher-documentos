from tkinter import *
import os
from docx import Document
from docx2pdf import convert
from tkinter import filedialog
from tkinter import messagebox


def tela():
    pass  # Adicione sua lógica aqui

def escolher_tela():
    if opcao_selecionada.get() == "Formulário Notebook":
        # Chamar a função para preencher o formulário de notebook
        preencher_documento_notebook()
    elif opcao_selecionada.get() == "Formulário Headset":
        # Chamar a função para preencher o formulário de headset
        preencher_documento_headset()
    elif opcao_selecionada.get() == "Declaração Isenção DANFe":
        # Chamar a função para preencher o formulário de headset
        preencher_documento_DANFe()
    elif opcao_selecionada.get() == "Declaração Isenção ICMS":
        # Chamar a função para preencher o formulário de headset
        preencher_documento_ICMS()
    elif opcao_selecionada.get() == "Declaração Isenção NF":
        # Chamar a função para preencher o formulário de headset
        preencher_documento_NF()

def preencher_documento(campos, nome_arquivo):
    # Carregar o modelo do documento
    doc = Document(nome_arquivo)
   
    # Iterar sobre o dicionário e substituir os campos no documento
    for campo, entrada in campos.items():
        if entrada:
            for p in doc.paragraphs:
                for run in p.runs:
                    if campo in run.text:
                        run.text = run.text.replace(campo, entrada)
 
    # Salvar o doc preenchido como docx temporário
    temp_docx = 'documento_preenchido.docx'
    doc.save(temp_docx)
 
    # Converter o documento docx para pdf
    convert(temp_docx)  # Isso criará um arquivo 'documento_preenchido.pdf'
 
    # Salvar o documento preenchido
    doc.save('documento_preenchido.docx')
 
def preencher_documento_notebook():
    import termo_notebook
    pass
 
def preencher_documento_headset():
    # Importar e executar o código para preencher o documento do headset
    import termo_headset
    pass

def preencher_documento_DANFe():
    # Importar e executar o código para preencher o documento do headset
    import isencao_DANFe
    pass

def preencher_documento_ICMS():
    # Importar e executar o código para preencher o documento do headset
    import isencao_ICMS
    pass

def preencher_documento_NF():
    # Importar e executar o código para preencher o documento do headset
    import isencao_NF
    pass
 
 
import shutil

def preencher_e_salvar():
    if opcao_selecionada.get() == "Formulário Notebook":
        preencher_documento_notebook()
    elif opcao_selecionada.get() == "Formulário Headset":
        preencher_documento_headset()
    elif opcao_selecionada.get() == "Declaração Isenção DANFe":
        preencher_documento_DANFe()
    elif opcao_selecionada.get() == "Declaração Isenção ICMS":
        preencher_documento_ICMS()
    elif opcao_selecionada.get() == "Declaração Isenção ICMS":
        preencher_documento_NF()
 
    # Mover o arquivo PDF para o diretório desejado
    novo_nome = 'documento_preenchido.pdf'
    novo_caminho_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=novo_nome, filetypes=[("PDF files", "*.pdf")])
    if novo_caminho_pdf:
        shutil.move('documento_preenchido.pdf', novo_caminho_pdf)
 

# Criar a janela principal
root = Tk()
root.title("Preencher Documento")

# Lista de opções para o formulário
opcoes_formulario = ["Formulário Notebook", "Formulário Headset", "Declaração Isenção DANFe", "Declaração Isenção ICMS", "Declaração Isenção NF"]
opcao_selecionada = StringVar(root)
opcao_selecionada.set(opcoes_formulario[0])  # Definir a primeira opção como padrão

# Criar o menu de opções para selecionar o formulário
formulario_menu = OptionMenu(root, opcao_selecionada, *opcoes_formulario)
formulario_menu.grid(row=15, column=0, columnspan=2, pady=10)

# Adicione aqui os rótulos e campos restantes para o formulário

# Botão para escolher o formulário
Button(root, text="Escolher Formulário", command=escolher_tela).grid(row=17, column=0, columnspan=2)

# Iniciar o loop principal da interface gráfica
root.mainloop()
