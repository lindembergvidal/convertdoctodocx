## 📄 DOC → DOCX Converter (Microsoft Word Automation)

Este projeto automatiza a conversão de arquivos .doc (Word binário legado) para .docx, utilizando o Microsoft Word via automação COM.
O resultado é idêntico ao “Salvar como” do Word, preservando hyperlinks embutidos, formatação e estrutura do documento.

⚠️ Esta é a única abordagem 100% fiel para documentos .doc corporativos e antigos.

## ✅ O que este script faz

Converte arquivos .doc → .docx

Mantém:

    Hyperlinks embutidos

    Campos antigos do Word

    Estrutura e formatação

    Processa todos os arquivos de uma pasta automaticamente

    Não utiliza LibreOffice, Google Drive ou ferramentas online

❌ O que ele não faz

    Não roda em Linux ou macOS

    Não funciona sem o Microsoft Word instalado

    Não converte arquivos protegidos por senha

## 🧰 Requisitos

    Windows

    Microsoft Word instalado e licenciado

    Python 3.8+

    VS Code ou terminal

    Biblioteca pywin32

## 📦 Instalação

No terminal (PowerShell ou VS Code):

    pip install pywin32

## 📁 Estrutura do projeto

Os arquivos ficam diretamente nas pastas do projeto:

project/
 ├── import/
 │    ├── teste.doc
 │    └── outro_arquivo.doc
 ├── output/
 │    ├── teste.docx
 │    └── outro_arquivo.docx
 └── doc_to_docx.py


import/ → coloque aqui os arquivos .doc

output/ → os .docx convertidos serão gerados aqui

## ▶️ Como usar

Coloque os arquivos .doc na pasta import/

Execute o script:

python doc_to_docx.py


Os arquivos convertidos aparecerão em output/

## 🧠 Como funciona (resumo técnico)

O script:

Abre o Microsoft Word em modo invisível

Abre cada arquivo .doc

Executa o equivalente a “Salvar como → DOCX”

Fecha o documento e segue para o próximo

Ou seja: é o próprio Word fazendo a conversão.

## 🔒 Por que essa abordagem é necessária?

Arquivos .doc utilizam um formato binário legado, que:

Não é totalmente suportado por bibliotecas open-source

Não é convertido corretamente por APIs cloud

Falha em ferramentas como Pandoc e Google Drive API

Por isso, a automação do Word é o padrão ouro para esse tipo de conversão.

## ⚠️ Observações importantes

O Word não deve estar aberto durante a execução

O script converte os arquivos sequencialmente

Documentos muito grandes podem levar alguns segundos

## 📜 Licença

Uso livre para fins educacionais e internos.
Verifique a licença do Microsoft Word para uso comercial.