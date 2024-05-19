# Automação de Preencher Documentos

## Descrição
Este projeto fornece uma interface gráfica para preencher automaticamente campos em um documento Word (.docx) e convertê-lo para PDF. É útil para automatizar a criação de documentos com informações repetitivas.

## Requisitos
Antes de executar o código, certifique-se de ter o Python instalado (qualquer versão).

## Instalação
Para instalar as dependências necessárias, siga os passos abaixo:

Abra o prompt de comando.
Navegue até o diretório onde o arquivo requirements.txt está localizado.
Execute o seguinte comando:

pip install -r requirements.txt

## Utilização
Execute o script principal.
Preencha os campos necessários na interface gráfica.
Clique em "Preencher e Salvar" para gerar o documento preenchido e convertido para PDF.

## Dependências
As principais dependências necessárias para este projeto são:

python-docx==0.8.11 (Para manipulação de arquivos .docx)
docx2pdf==0.1.8 (Para conversão de arquivos .docx para .pdf)
Certifique-se de que estas dependências estejam incluídas no seu arquivo requirements.txt.

## Estrutura do Projeto
Abaixo está um exemplo de estrutura do projeto:


## Automação-de-Preencher-Documentos/
│
├── index.py
├── termo_notebook.py
├── termo_notebook.docx
├── termo_headset.py
├── termo_headset.docx
├── isencao_NF.py
├── isencao_ICMS.py
├── isencao_DANFe.py
│
├── Declaração isenção DANFe - DHL preenchido.docx
├── declaracao_ICMS - preenchido.docx
├── declaracao_ICMS.docx
├── declaracao_NF - preenchido.docx
├── declaracao_NF.docx
├── declaração isencao DANFe - DHL.docx
├── headset_preenchido.docx
├── termo notebook_preenchido.docx
│
└── README.md

## Autor
Lucas Kevin
Contato: lucaskevin455@gmail.com
