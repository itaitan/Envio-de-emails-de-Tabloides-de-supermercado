# instalação do python site = https://www.python.org/downloads/ versão 3.9.0
# Verificar no cmd Windowns cod - python --version (instalado corretamente)

import smtplib  # Não precisa instalar ja vem no python
# Não precisa instalar ja vem no python
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText  # Não precisa instalar ja vem no python
from email.mime.base import MIMEBase  # Não precisa instalar ja vem no python
import xlrd  # instalar no terminal pip3 install xlrd==1.2.0 boto3
import pdb  # Não precisa instalar ja vem no python
import time  # Não precisa instalar ja vem no python
import email.utils  # Não precisa instalar ja vem no python
from datetime import datetime  # instalar no terminal pip3 install datetime

print("Iniciando...")

arq = open("log.txt", "w")

texto = list()

fromaddr = "contato@enviodeemails.tk"  # Email Remetente
remetente = "COMERCIAL ESPERANÇA"  # Nome Remetente
senha = "ws2X3%m7"
SSLEnd = 'mail.enviodeemails.tk:'
SSLPort = str(465)
SSL = SSLEnd + SSLPort

msg = MIMEMultipart()
msg['From'] = email.utils.formataddr((remetente, fromaddr))
# Assunto do Email
msg['Subject'] = "Economia de variedade de produtos? Aqui no Comercial Esperança você encontra!"
# Processo do trabalho, O Designer enviar o HTML recortado ja hospedado! em arquivo, Abrindo o mesmo em
# Chrome Clickar em inspecionar abrindo o CMD do Chrome clicando no codigo HTML <Body> com botão direito
# Copy > CopyOtherHTML que seria o código que esta abaixo você vai colar do Body para baixo apenas.
corpo_email = """<!doctype html><html>
<head>
<title>Economia de variedade de produtos? Aqui no Comercial Esperança você encontra!</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body bgcolor="#F6F6F6" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table id="Tabela_01" width="800" height="1419" border="0" cellpadding="0" cellspacing="0" align="center">
	<tbody><tr>
		<td colspan="4">
			<a href="http://comercialesperanca.net.br/ofertas/" target="_blank">
				<img src="https://i.ibb.co/wrt1h4T/head.png" width="800" height="145" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
	</tr>
	<tr>
		<td colspan="4">
			<a href="http://comercialesperanca.net.br/ofertas/" target="_blank">
				<img src="https://i.ibb.co/3Tp4B2s/of.png" width="800" height="991" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
	</tr>
	<tr>
		<td>
			<a href="http://comercialesperanca.net.br/ofertas/" target="_blank">
				<img src="https://i.ibb.co/qkhHb0N/data.png" width="516" height="80" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
		<td>
			<a href="https://www.facebook.com/ComlEsperanca/" target="_blank">
				<img src="https://i.ibb.co/z6YbQJC/fb.png" width="84" height="80" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
		<td>
			<a href="https://www.instagram.com/comercialesperanca.oficial/" target="_blank">
				<img src="https://i.ibb.co/rk30Bf2/ig.png" width="93" height="80" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
		<td>
			<a href="https://www.linkedin.com/company/comercialesperanca/" target="_blank">
				<img src="https://i.ibb.co/XkWTYwm/in.jpg" width="107" height="80" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
	</tr>
	<tr>
		<td colspan="4">
			<a href="http://comercialesperanca.net.br/ofertas/" target="_blank">
				<img src="https://i.ibb.co/mDKVbZw/baixo.png" width="800" height="203" border="0" style="display:block; margin:0; padding:0; border:none;"></a></td>
	</tr>
</tbody></table>

</body>
</html>"""

part1 = MIMEText(corpo_email, "html")
msg.attach(part1)

s = smtplib.SMTP_SSL(SSL)

s.login(fromaddr, senha)

text = msg.as_string()

emails = []
# Ler do Excel:
workbook = xlrd.open_workbook('mailing.xlsx')
sheet = workbook.sheet_by_index(0)

# Ler somente linhas, uma vez que não temos mais de 1 coluna
# No excel funciona a orientação como Matriz, sendo linha,coluna
# Nesse caso, linha,0, ou seja, na 1a coluna
# Os indices começam do 0
# Quantidade linha do excell && caso de erro, verificar no log qual for email e modificar o numero da onde ele parou!
# log é o arquivo TXT da mesma pasta
for linha in range(0, 4):
    for coluna in range(0, 1):  # Aqui é a primeira coluna
        emails.append(sheet.cell_value(linha, coluna))
# Escrevendo o log dos emails passado,
for email in emails:
    agora = datetime.now()
    agora_string = agora.strftime("%H:%M:%S")
    toaddr = email
    s.sendmail(fromaddr, toaddr, text)
    texto = str(toaddr) + ' ' + agora_string
    arq.write(texto + '\n')


s.quit()
print('############################################################')
msg = "Os emails foram enviados com sucesso!"
print(msg)
arq.write('\n' + msg)

arq.close()
