import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from helpers import getMessageContent,getnameAttachemnt,getccmail
text = "Inventario del día"
html = getMessageContent()
html = html.replace("{{dia}}","2028-08-01")
text_part = MIMEText(text, 'plain')
html_part = MIMEText(html, 'html')

msg_alternative = MIMEMultipart('alternative')
msg_alternative.attach(text_part)
msg_alternative.attach(html_part)

filename= getnameAttachemnt()
pathadjunto = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents','ReporteadorInventario',filename)
fp=open(pathadjunto,'rb')
attachment = MIMEApplication(fp.read(),_subtype="xlsx")
fp.close()
attachment.add_header('Content-Disposition', 'attachment', filename=filename)

msg_mixed = MIMEMultipart('mixed')
msg_mixed.attach(msg_alternative)
msg_mixed.attach(attachment)
msg_mixed['From'] = 'almacensat@coronalostuxtlas.com.mx'
destinatario = ["auditoriasistemas@coronalostuxtlas.com.mx"]
cc = getccmail()
msg_mixed['To'] = ",".join(destinatario)
msg_mixed['Cc'] = ",".join(cc)
msg_mixed['Subject'] = 'Inventario Prueba Sistemas'

smtp_obj = smtplib.SMTP_SSL('smtp.coronalostuxtlas.com.mx')
smtp_obj.ehlo()
smtp_obj.login('almacensat@coronalostuxtlas.com.mx', 'Alm$sat&22')
smtp_obj.sendmail(msg_mixed['From'], (destinatario+cc), msg_mixed.as_string())
smtp_obj.quit()