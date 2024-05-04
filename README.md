# Automação do WhatsApp com Python

Este projeto é um script Python que automatiza o envio de mensagens pelo WhatsApp. Ele lê os contatos de uma base de dados inseridos em uma planilha do Excel e envia uma mensagem personalizada para cada contato.

## Dependências

O script depende das seguintes bibliotecas Python:

- openpyxl
- urllib
- webbrowser
- time
- pyautogui

## Como usar

1. Abra o WhatsApp Web no navegador.
2. Aguarde 15 segundos para que a página seja carregada.
3. O script lê os dados dos contatos da planilha 'contatos.xlsx' e da página 'Planilha1', logo é necessário preenchelo com informações válidas.
4- Para preencher a planilha de dados informe na Coluna A(n) o nome do usuário, logo em seguida, na coluna B(n) informe o núemro de telefone seguindo o formato (<DDD>) <sequência de digitos>.
5. Para cada contato na planilha, o script lê o nome e o número de telefone.
6. Uma mensagem personalizada é criada para cada contato.
7. O script tenta enviar a mensagem através do WhatsApp Web.
8. Se a mensagem não puder ser enviada, o nome e o número de telefone do contato são gravados em 'erros.csv'.
