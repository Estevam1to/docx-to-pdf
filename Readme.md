# Conversor DOCX para PDF

Este é um conversor de arquivos DOCX para PDF escrito em Rust. Ele permite converter documentos do Microsoft Word (.docx) para o formato PDF, mantendo a formatação básica, incluindo texto, imagens e tabelas simples.

## Funcionalidades

- Conversão de arquivos DOCX para PDF
- Suporte para texto, imagens e tabelas simples
- Manutenção de formatação básica
- Redimensionamento e centralização de imagens
- Logging para acompanhamento do processo de conversão

## Pré-requisitos

- Rust (versão estável mais recente)
- Cargo (geralmente vem com a instalação do Rust)

## Instalação

1. Clone este repositório:
   ```
   git clone https://github.com/Estevam1to/docx-to-pdf
   cd docx-to-pdf
   ```

2. Compile o projeto:
   ```
   cargo build --release
   ```

## Uso

Execute o programa a partir da linha de comando, fornecendo o arquivo DOCX de entrada e o nome desejado para o arquivo PDF de saída:

### Exemplo detalhado

Suponha que você tenha um arquivo DOCX chamado "innput.docx" e queira convertê-lo para PDF:

1. Abra o terminal e navegue até o diretório do projeto.

2. Execute o comando:
   ```
   cargo run input.docx output.pdf
   ```