import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# Ler o arquivo Excel com os IDs e e-mails dos colaboradores
df = pd.read_excel('/home/pdfconvert/cadastro-funcioario.xlsx', dtype={'ID': str})

# Diretório onde estão as pastas dos colaboradores
diretorio_base = '/home/pdfconvert/Output'

# Função para encontrar o PDF mais recente em uma pasta
def encontrar_pdf_mais_recente(id_colaborador):
    for pasta in os.listdir(diretorio_base):
        if pasta.startswith(id_colaborador):
            diretorio_colaborador = os.path.join(diretorio_base, pasta)
            arquivos = os.listdir(diretorio_colaborador)
            arquivos_pdf = [arquivo for arquivo in arquivos if arquivo.lower().endswith('.pdf')]
            if arquivos_pdf:
                arquivos_pdf.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio_colaborador, x)), reverse=True)
                return os.path.join(diretorio_colaborador, arquivos_pdf[0])
                 
    return None

# Configurações do servidor de e-mail
email_host = 'smtp-mail.outlook.com'
email_porta = 587
email_usuario = 'notify@bmj.com.br'
email_senha = '35t@980197!!'

# Função para enviar e-mail com o PDF anexado
def enviar_email(destinatario, anexo):
    mensagem = MIMEMultipart()
    mensagem['From'] = email_usuario
    mensagem['To'] = destinatario
    mensagem['Subject'] = 'Contracheque'

    corpo_mensagem = 'Olá,\n\nSegue em anexo contracheque.\n\nAtenciosamente,\nSetor Financeiro.\n Por favor, não responda a este e-mail. Em caso de dúvidas, entre em contato com: financeiro@bmj.com.br'
    mensagem.attach(MIMEText(corpo_mensagem, 'plain'))

    arquivo_anexo = open(anexo, 'rb')
    anexo_mime = MIMEBase('application', 'octet-stream')
    anexo_mime.set_payload(arquivo_anexo.read())
    encoders.encode_base64(anexo_mime)
    anexo_mime.add_header('Content-Disposition', f'attachment; filename={os.path.basename(anexo)}')
    mensagem.attach(anexo_mime)

    servidor_smtp = smtplib.SMTP(email_host, email_porta)
    servidor_smtp.starttls()
    servidor_smtp.login(email_usuario, email_senha)
    texto_email = mensagem.as_string()
    servidor_smtp.sendmail(email_usuario, destinatario, texto_email)
    servidor_smtp.quit()

# Iterar sobre os colaboradores e enviar o PDF mais recente por e-mail
for indice, linha in df.iterrows():
    id_colaborador = linha['ID']
    pdf_recente = encontrar_pdf_mais_recente(id_colaborador)
    if pdf_recente:
        email_destinatario = linha['E-mail']
        enviar_email(email_destinatario, pdf_recente)
        print(f'E-mail enviado para {email_destinatario} com ID={id_colaborador}')
    else:
        print(f'Nenhum PDF encontrado para ID {id_colaborador}')