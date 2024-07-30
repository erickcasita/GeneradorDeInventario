import datetime, smtplib,time,os
from progress1bar import ProgressBar
from email.message import EmailMessage
from email.utils import formataddr
def validatedate(date_text):
        try:
            datetime.date.fromisoformat(date_text)
            return True
        except ValueError:
          print ("\n Formato de fecha incorrecto, Formato:  YYYY-MM-DD")
          
def getMessageContent():
  with open('mails/MailTemplate.html') as fichero:
    
    return fichero.read()
def getnameAttachemnt():
  with open('mails/attachment.name.mail') as fichero:
    
    return fichero.read() 
  
def getccmail():
  cc = [] 
  with open('mails/mails.em', 'r') as fichero:
    for linea in fichero:
      linea = linea.replace('\n','')
      cc.append(linea)
  return cc
def sendMail(date, remitente,destinatario,cc):
  #NameAttachment
  adjuntoname= getnameAttachemnt()
  pathadjunto = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents','ReporteadorInventario',adjuntoname)
  mensaje = getMessageContent()
  mensaje = mensaje.replace("{{dia}}", date)
  email = EmailMessage()
  email["From"] = formataddr(('Almacén SAT | Corona los Tuxtlas ', remitente))
  email["To"] = destinatario
  email["Cc"] = cc
  email["Subject"] = "INVENTARIOS DE ALMACENES "+date + " 🍻"
  email.set_content(mensaje, subtype="html")
  with open(pathadjunto, "rb") as f:
    email.add_attachment(
        f.read(),
        filename=adjuntoname,
        maintype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        subtype="xlsx"
    )
  smtp = smtplib.SMTP_SSL("smtp.coronalostuxtlas.com.mx")
  smtp.login(remitente, "Alm$sat&22")
  smtp.sendmail(remitente,(destinatario + cc), email.as_string())
  smtp.quit()
  #Remove Attachment.name.file
  os.remove('mails/attachment.name.mail')
def progressbarmail():
  kwargs = {
    'total': 100,
    'completed_message': 'Proceso Terminado',
    'clear_alias': False,
    'show_fraction': False,
    'show_prefix': False,
    'show_duration': True
}
  with ProgressBar(**kwargs) as pb:
      pb.alias = 'Envío de  correro'
      for _ in range(pb.total):
          pb.count += 1
          time.sleep(0.3)
