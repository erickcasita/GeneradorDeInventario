import datetime, smtplib,time,os,locale
from progress1bar import ProgressBar
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
def validatedate(date_text):
        try:
            datetime.date.fromisoformat(date_text)
            return True
        except ValueError:
          print ("\n Formato de fecha incorrecto, Formato:  YYYY-MM-DD")
def getnameAttachemnt():
  with open('mails/attachment.name.mail') as fichero:
    return fichero.read()
def getmailcontent():
    with open('mails/mailcontent.em', 'r',encoding='utf-8') as fichero:
      return fichero.read()
def getemailto():
  to = []
  with open('mails/mails.em', 'r') as fichero:
    for linea in fichero:
      linea = linea.replace('\n','')
      to.append(linea)
  return to
def sendMail():
  to = getemailto()
  total = len(to)
  kwargs = {
    'total': total,
    'completed_message': 'Proceso Terminado',
    'clear_alias': False,
    'show_fraction': False,
    'show_prefix': False,
    'show_duration': True
}
  with ProgressBar(**kwargs) as pb:
      for emailto in to:
        pb.alias = 'Envío de correo'
        destinatario = []
        destinatario.append(emailto)
        locale.setlocale(locale.LC_ALL, 'es_ES.utf8')
        date = datetime.datetime.strftime(datetime.datetime.now(),'%A %d de %B del %Y')
        text = "Inventario del día"
        html = getmailcontent()
        html = html.replace("{{dia}}",date)
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
        msg_mixed['To'] = ",".join(destinatario)
        msg_mixed['Subject'] = 'Inventario de almacenes al dia ' + str(str(datetime.datetime.strftime(datetime.datetime.now(),'%A %d de %B del %Y')))
        smtp_obj = smtplib.SMTP_SSL('smtp.coronalostuxtlas.com.mx')
        smtp_obj.ehlo()
        smtp_obj.login('almacensat@coronalostuxtlas.com.mx', 'Alm$sat&22')
        smtp_obj.sendmail(msg_mixed['From'], (destinatario), msg_mixed.as_string())
        smtp_obj.quit()        
        time.sleep(1)
        pb.count += 1
      os.remove('mails/attachment.name.mail')