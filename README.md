📊 Automação de Relatórios Diários
Este projeto é um script em Python que automatiza o processo de login em um sistema web, exportação de relatórios em Excel e envio desses relatórios por e-mail para cada loja cadastrada.

🚀 Funcionalidades
Login automático no sistema via navegador.

Exportação de relatórios em formato .xlsx usando automação de tela.

Tratamento dos arquivos Excel:

Remove colunas desnecessárias.

Ajusta cabeçalhos e formatação.

Congela a primeira linha para facilitar leitura.

Envio automático por e-mail via Outlook, com anexos e assunto personalizado.

Leitura de lojas e credenciais a partir de um arquivo lojas.csv.

📂 Estrutura do Projeto
main.py → Script principal com toda a lógica.

lojas.csv → Arquivo contendo lista de lojas, usuários e senhas.

Pasta de Downloads → Local onde os relatórios são baixados.

Pasta de Relatórios → Local onde os arquivos tratados são salvos.

⚙️ Configurações
No início do código existem variáveis que podem ser ajustadas conforme seu ambiente:

URL_LOGIN → URL da página de login do sistema.

DOWNLOADS_DIR → Diretório padrão de downloads do navegador.

DESTINO_DIR → Diretório onde os relatórios tratados serão salvos.

emails_por_loja → Dicionário com os e-mails de cada loja.

alias_lojas → Mapeamento de nomes alternativos para padronização.

📑 Estrutura do CSV (lojas.csv)
O arquivo deve conter colunas como:

csv
loja,usuario,senha
Loja A,usuarioA,senhaA
Loja B,usuarioB,senhaB
Loja C,usuarioC,senhaC
🖥️ Dependências
O script utiliza as seguintes bibliotecas:

os, time, glob, csv, datetime → Bibliotecas padrão do Python.

webbrowser → Para abrir o navegador.

pyautogui → Automação de cliques e teclado.

openpyxl → Manipulação de arquivos Excel.

win32com.client → Integração com Outlook para envio de e-mails.

Instale as dependências com:

bash
pip install pyautogui openpyxl pywin32
▶️ Como Executar
Configure o arquivo lojas.csv com as lojas e credenciais.

Ajuste os diretórios (DOWNLOADS_DIR e DESTINO_DIR) conforme seu ambiente.

Execute o script:

bash
python main.py
O programa irá:

Abrir o navegador e fazer login.

Exportar relatórios de cada loja.

Tratar os arquivos Excel.

Enviar os relatórios por e-mail automaticamente.

⚠️ Observações
As coordenadas de tela (COORD_USUARIO, COORD_SENHA, etc.) foram configuradas para resolução 1366x768. Se sua tela tiver outra resolução, ajuste os valores.

O envio de e-mails depende de ter o Outlook instalado e configurado na máquina.

Este código é um template genérico. Substitua os valores fictícios (URLs, e-mails, credenciais) pelos dados reais do seu ambiente.
