import datetime
def validatedate(date_text):
        try:
            datetime.date.fromisoformat(date_text)
            return True
        except ValueError:
          print ("\n Formato de fecha incorrecto, Formato:  YYYY-MM-DD")
          
def getMessageContent():
  with open('mails/MailTemplate.html') as fichero:
    
    return fichero.read()
