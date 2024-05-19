<title>README - Automa-o-de-Preencher-Documentos</title>

<h1>Automa-o-de-Preencher-Documentos</h1>

    <h2>Descrição</h2>

    <p>Este projeto fornece uma interface gráfica para preencher automaticamente campos em um documento Word (.docx) e
        convertê-lo para PDF. É útil para automatizar a criação de documentos com informações repetitivas.</p>

    <h2>Requisitos</h2>

    <p>Antes de executar o código, certifique-se de ter o Python instalado (qualquer versão).</p>

    <h2>Instalação</h2>

    <p>Para instalar as dependências necessárias, siga os passos abaixo:</p>

    <ol>
        <li>Abra o prompt de comando.</li>
        <li>Navegue até o diretório onde o arquivo <code>requirements.txt</code> está localizado.</li>
        <li>Execute o seguinte comando:</li>
    </ol>

    <pre><code>pip install -r requirements.txt</code></pre>

    <h2>Utilização</h2>

    <ol>
        <li>Execute o script principal.</li>
        <li>Preencha os campos necessários na interface gráfica.</li>
        <li>Clique em "Preencher e Salvar" para gerar o documento preenchido e convertido para PDF.</li>
    </ol>

    <h2>Dependências</h2>

    <p>As principais dependências necessárias para este projeto são:</p>

    <ul>
        <li><code>python-docx==0.8.11</code> (Para manipulação de arquivos .docx)</li>
        <li><code>docx2pdf==0.1.8</code> (Para conversão de arquivos .docx para .pdf)</li>
    </ul>

    <p>Certifique-se de que estas dependências estejam incluídas no seu arquivo <code>requirements.txt</code>.</p>

    <h2>Estrutura do Projeto</h2>

    <p>Abaixo está um exemplo de estrutura do projeto:</p>

    <pre>
        <code>Automa-o-de-Preencher-Documentos/
        ├── declaracao_ICMS.docx  # Modelo de documento
        ├── preenchimento_doc.py  # Script principal
        ├── requirements.txt      # Arquivo de dependências
        └── README.md             # Este arquivo
        </code>
    </pre>

    <h2>Autor</h2>

    <ul>
        <li>Seu Nome</li>
        <li>Contato: seu-email@example.com</li>
    </ul>
