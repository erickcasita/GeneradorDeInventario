import datetime, smtplib,time,os,locale, email
from progress1bar import ProgressBar
from email import encoders
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
def validatedate(date_text):
        try:
            datetime.date.fromisoformat(date_text)
            return True
        except ValueError:
          print ("\n Formato de fecha incorrecto, Formato:  YYYY-MM-DD")
          
def getMessageContent():
  html = """ <!doctype html>
<html lang="es">
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <title>Email Agencia Corona</title>
    <style media="all" type="text/css">
    /* -------------------------------------
    GLOBAL RESETS
------------------------------------- */
    
    body {
      font-family: Helvetica, sans-serif;
      -webkit-font-smoothing: antialiased;
      font-size: 16px;
      line-height: 1.3;
      -ms-text-size-adjust: 100%;
      -webkit-text-size-adjust: 100%;
    }
    
    table {
      border-collapse: separate;
      mso-table-lspace: 0pt;
      mso-table-rspace: 0pt;
      width: 100%;
    }
    
    table td {
      font-family: Helvetica, sans-serif;
      font-size: 16px;
      vertical-align: top;
    }
    /* -------------------------------------
    BODY & CONTAINER
------------------------------------- */
    
    body {
      background-color: #f4f5f6;
      margin: 0;
      padding: 0;
    }
    
    .body {
      background-color: #f4f5f6;
      width: 100%;
    }
    
    .container {
      margin: 0 auto !important;
      max-width: 600px;
      padding: 0;
      padding-top: 24px;
      width: 600px;
    }
    
    .content {
      box-sizing: border-box;
      display: block;
      margin: 0 auto;
      max-width: 600px;
      padding: 0;
    }
    /* -------------------------------------
    HEADER, FOOTER, MAIN
------------------------------------- */
    
    .main {
      background: #ffffff;
      border: 1px solid #eaebed;
      border-radius: 16px;
      width: 100%;
    }
    
    .wrapper {
      box-sizing: border-box;
      padding: 24px;
    }
    
    .footer {
      clear: both;
      padding-top: 24px;
      text-align: center;
      width: 100%;
    }
    
    .footer td,
    .footer p,
    .footer span,
    .footer a {
      color: #9a9ea6;
      font-size: 16px;
      text-align: center;
    }
    /* -------------------------------------
    TYPOGRAPHY
------------------------------------- */
    
    p {
      font-family: Helvetica, sans-serif;
      font-size: 16px;
      font-weight: normal;
      margin: 0;
      margin-bottom: 16px;
    }
    
    a {
      color: #0867ec;
      text-decoration: underline;
    }
    /* -------------------------------------
    BUTTONS
------------------------------------- */
    
    .btn {
      box-sizing: border-box;
      min-width: 100% !important;
      width: 100%;
    }
    
    .btn > tbody > tr > td {
      padding-bottom: 16px;
    }
    
    .btn table {
      width: auto;
    }
    
    .btn table td {
      background-color: #ffffff;
      border-radius: 4px;
      text-align: center;
    }
    
    .btn a {
      background-color: #ffffff;
      border: solid 2px #0867ec;
      border-radius: 4px;
      box-sizing: border-box;
      color: #0867ec;
      cursor: pointer;
      display: inline-block;
      font-size: 16px;
      font-weight: bold;
      margin: 0;
      padding: 12px 24px;
      text-decoration: none;
      text-transform: capitalize;
    }
    
    .btn-primary table td {
      background-color: #0867ec;
    }
    
    .btn-primary a {
      background-color: #0867ec;
      border-color: #0867ec;
      color: #ffffff;
    }
    
    @media all {
      .btn-primary table td:hover {
        background-color: #ec0867 !important;
      }
      .btn-primary a:hover {
        background-color: #ec0867 !important;
        border-color: #ec0867 !important;
      }
    }
    
    /* -------------------------------------
    OTHER STYLES THAT MIGHT BE USEFUL
------------------------------------- */
    
    .last {
      margin-bottom: 0;
    }
    
    .first {
      margin-top: 0;
    }
    
    .align-center {
      text-align: center;
    }
    
    .align-right {
      text-align: right;
    }
    
    .align-left {
      text-align: left;
    }
    
    .text-link {
      color: #0867ec !important;
      text-decoration: underline !important;
    }
    
    .clear {
      clear: both;
    }
    
    .mt0 {
      margin-top: 0;
    }
    
    .mb0 {
      margin-bottom: 0;
    }
    
    .preheader {
      color: transparent;
      display: none;
      height: 0;
      max-height: 0;
      max-width: 0;
      opacity: 0;
      overflow: hidden;
      mso-hide: all;
      visibility: hidden;
      width: 0;
    }
    
    .powered-by a {
      text-decoration: none;
    }
    
    /* -------------------------------------
    RESPONSIVE AND MOBILE FRIENDLY STYLES
------------------------------------- */
    
    @media only screen and (max-width: 640px) {
      .main p,
      .main td,
      .main span {
        font-size: 16px !important;
      }
      .wrapper {
        padding: 8px !important;
      }
      .content {
        padding: 0 !important;
      }
      .container {
        padding: 0 !important;
        padding-top: 8px !important;
        width: 100% !important;
      }
      .main {
        border-left-width: 0 !important;
        border-radius: 0 !important;
        border-right-width: 0 !important;
      }
      .btn table {
        max-width: 100% !important;
        width: 100% !important;
      }
      .btn a {
        font-size: 16px !important;
        max-width: 100% !important;
        width: 100% !important;
      }
    }
    /* -------------------------------------
    PRESERVE THESE STYLES IN THE HEAD
------------------------------------- */
    
    @media all {
      .ExternalClass {
        width: 100%;
      }
      .ExternalClass,
      .ExternalClass p,
      .ExternalClass span,
      .ExternalClass font,
      .ExternalClass td,
      .ExternalClass div {
        line-height: 100%;
      }
      .apple-link a {
        color: inherit !important;
        font-family: inherit !important;
        font-size: inherit !important;
        font-weight: inherit !important;
        line-height: inherit !important;
        text-decoration: none !important;
      }
      #MessageViewBody a {
        color: inherit;
        text-decoration: none;
        font-size: inherit;
        font-family: inherit;
        font-weight: inherit;
        line-height: inherit;
      }
    }
    </style>
  </head>
  <body>
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="body">
      <tr>
        <td>&nbsp;</td>
        <td class="container">
          <div class="content">

            <!-- START CENTERED WHITE CONTAINER -->
            <span class="preheader">Inventario de almacenes al periodo {{dia}}.</span>
            <table role="presentation" border="0" cellpadding="0" cellspacing="0" class="main">

              <!-- START MAIN CONTENT AREA -->
              <tr>
                <td class="wrapper">
                  <img src="https://i.imgur.com/X2mQfjW.png" alt="Logo - Corona" width="300" height="260" border="0" style="border:0; outline:none; text-decoration:none; display:block; margin-left:auto ; margin-right: auto;">
                  <div style="font-family:'Helvetica Neue',Arial,sans-serif;font-size:28px;font-weight:bold;line-height:1;text-align:center;color:#555;">
                    Inventario de Almacenes al dia {{dia}}
                  </div>
                  <br>
                  <p>¡Buen día!</p>
                  <br>
                  <p>Por medio de la presente, les hago llegar el inventario correspondiente al almacén de lleno San andrés Tuxlta y almacén de lleno Juan Díaz Covarrubias. </p>
                  <br>
                  <p>Saludos Cordiales</p>
                </td>
              </tr>
              

              <!-- END MAIN CONTENT AREA -->
              </table>

            <!-- START FOOTER -->
            <div class="footer">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td class="content-block">
                    <span class="apple-link">Comercializadora de Cervezas de los Tuxtlas S.A DE C.V. </span>
                    <br> Encargado de almacén -- almacensat@coronalostuxtlas.com.mx
                  </td>
                </tr>
                <tr>
                  <td class="content-block powered-by">
                    Desarrollado por  <a href="#">Ing. Erick M. Ramírez Casanova</a>
                  </td>
                </tr>
              </table>
            </div>

            <!-- END FOOTER -->
            
<!-- END CENTERED WHITE CONTAINER --></div>
        </td>
        <td>&nbsp;</td>
      </tr>
    </table>
  </body>
</html> """
  return html 
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
def sendMail():
  #date = datetime.datetime.strftime(datetime.datetime.now(),'%d-%m-%Y')
  with open('mails/mails.em', 'r') as fichero:
    for linea in fichero:
      to = []
      linea = linea.replace('\n','');
      to.append(linea)
      locale.setlocale(locale.LC_ALL, 'es_ES.utf8')
      date = datetime.datetime.strftime(datetime.datetime.now(),'%A %d de %B del %Y')
      text = "Inventario del día"
      html = getMessageContent()
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
      destinatario = to
      msg_mixed['To'] = ",".join(destinatario)
      msg_mixed['Subject'] = 'Inventario de almacenes al dia ' + str(str(datetime.datetime.strftime(datetime.datetime.now(),'%A %d de %B del %Y')))

      smtp_obj = smtplib.SMTP_SSL('smtp.coronalostuxtlas.com.mx')
      smtp_obj.ehlo()
      smtp_obj.login('almacensat@coronalostuxtlas.com.mx', 'Alm$sat&22')
      smtp_obj.sendmail(msg_mixed['From'], (destinatario), msg_mixed.as_string())
      smtp_obj.quit()
      time.sleep(1)
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
      pb.alias = 'Envío de correo'
      for _ in range(pb.total):
          pb.count += 1
          time.sleep(0.5)


getccmail()