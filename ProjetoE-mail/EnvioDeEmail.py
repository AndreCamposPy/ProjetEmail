#!/usr/bin/env python
# coding: utf-8

# In[1]:


# Importa os módulos necessários para enviar e-mails
import smtplib                                      # Módulo para comunicação com o servidor de e-mail
import ssl                                          # Módulo para criar um contexto seguro de conexão
from email.message import EmailMessage              # Classe para criar mensagens de e-mail
import os                                           # Módulo para interagir com o sistema de arquivos
import mimetypes                                    # Módulo para determinar o tipo MIME dos anexos


# In[2]:


# Carrega a senha do e-mail a partir de um arquivo separado
email_senha = open(r'C:\Users\Pichau\Desktop\ProjetoE-mail\keys\senha', 'r').read()

# Endereço de e-mail de origem e destino
email_origem = 'camposandre73679@gmail.com'  # Insira seu endereço de e-mail
email_destino = ('menezes.campos@yahoo.com')  # Insira o e-mail do destinatário


# # Mãos ao código!
# ### Vamos começar no ambiente do anaconda, onde vou inicialmente incorporar os módulos utilizados.

# ## Inserimos emails e senha
# ### Evite salvar senhas no corpo do código, assim, ao compartilhá-lo outras pessoas não terão acesso às suas senhas.
# ### Então, salvarei minha senha de email num arquivo .txt e usarei a função Open para gravar o conteúdo desse arquivo na varável "email_senha".

# ## Inserimos Assunto e Corpo do email
# ### Supondo que enviaremos um relatório atualizado sobre os mesmos indicadores, vou deixar o Assunto do email estático descrito no próprio código.
# ### Porém, pode ser que seja interessante mudar a mensagem do email com uma certa frequência, por isso vou deixá-la num arquivo externo onde poderei alterá-la futuramente sem a necessidade de editar o código em si.

# In[3]:


# Assunto e corpo do e-mail
assunto = 'Funil de Produção Poc 08-2023'
# Lê o corpo do e-mail a partir de um arquivo
body = open(r'C:\Users\Pichau\Desktop\ProjetoE-mail\corpo_email.txt', 'r').read() 


# ### Definição de configurações internas para envio do email.

# In[4]:


# Cria uma instância da classe EmailMessage para construir a mensagem de e-mail
mensagem = EmailMessage()

# Define quem está enviando o e-mail (endereço de e-mail de origem)
mensagem['From'] = email_origem

# Define o destinatário do e-mail (endereço de e-mail do destinatário)
mensagem['To'] = email_destino

# Define o assunto do e-mail
mensagem['Subject'] = assunto

# Define o conteúdo principal do e-mail (texto do corpo)
mensagem.set_content(body)


# ## Inserimos o caminho do(s) anexo(s)
# ### Usarei 2 anexos como exemplo: uma planilha excel e um gráfico.
# ### Crierei uma lista com os caminhos onde estão cada um dos anexos para envio.
# ### Você pode acrescentar à lista mais outros caminhos para enviar quantos anexos desejar.

# In[5]:


# Lista de caminhos dos arquivos a serem anexados
anexos = [r"C:\Users\Pichau\Desktop\ProjetoE-mail\imagem.jpg",
          r"C:\Users\Pichau\Desktop\ProjetoE-mail\base de dados.xlsx"]

# Adiciona cada anexo à mensagem
for anexo_path in anexos:
    mime_type, _ = mimetypes.guess_type(anexo_path)
    mime_type, mime_subtype = mime_type.split('/', 1)
    
    # Lê e anexa os arquivos ao e-mail
    with open(anexo_path, 'rb') as ap:
        mensagem.add_attachment(ap.read(), maintype=mime_type, subtype=mime_subtype,
                                filename=os.path.basename(anexo_path))


# # Etapa de envio do email
# #### Se o envio for concluído com sucesso, será exibida a mensagem "Email enviado com sucesso!!!"
# #### Se houver erro de autenticação (e-mail/senha incorretos) será exibida a mensagem "Email e/ou senha incorreto!"
# #### Caso ocorra algum outro erro será exibida a mensagem "Erro imprevisto"

# In[6]:


# O bloco 'try' é usado para envolver um código onde podem ocorrer exceções (erros).
try:
   
    # O código dentro deste bloco é executado normalmente até que uma exceção (ou erro) seja lançada.
    # Cria um contexto seguro para conexão SSL criptografando o conteúdo do email
    safe = ssl.create_default_context()

    # Conecta-se ao servidor SMTP seguro do Gmail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=safe) as smtp:
        # Realiza o login no e-mail de origem
        smtp.login(email_origem, email_senha)
        # Envia o e-mail
        smtp.sendmail(email_origem, email_destino, mensagem.as_string())
    print('Email enviado com sucesso!!!')

except smtplib.SMTPAuthenticationError:
    # Se ocorrer um erro de autenticação (e-mail/senha incorretos)
    print('Email e/ou senha incorreto!')

except Exception as erro:
    # Se ocorrer qualquer outro erro
    print('Erro imprevisto')
    # Imprime o tipo de erro que ocorreu
    print(f'O tipo do erro encontrado foi {erro.__class__}')
    # Imprime a descrição detalhada do erro
    print(f'O texto do erro encontrado foi {erro.__str__()}')


# # Fim.
