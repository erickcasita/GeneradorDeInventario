import requests,datetime

def Get_current_date():
    url = 'https://software-api-rnig.onrender.com/currentdate' #'http://localhost:3000/currentdate' 
    data = requests.get(url)
    try:
        if (data.status_code == 200):
            data  = data.json()
            return data['CurrentDate']
          
        else:
            return "error";
    except ValueError:
        print("Error al obtener la fecha del servidor");
        
def Get_current_expiration_software():
    url = 'https://software-api-rnig.onrender.com/currentversion'#'http://localhost:3000/currentversion' 
    data = requests.get(url)
    try:
        if (data.status_code == 200):
            data  = data.json()
            return data['Expiration']
          
        else:
            return "error";
    except ValueError:
        print("Error al obtener la fecha del servidor");

def Check_expiration_software(datecurrent, datesoftware):
    
    if (datecurrent !="error" and datesoftware != "error"):
        dateFormatter = "%Y-%m-%d"
        fecha_server=datetime.datetime.strptime(datecurrent, dateFormatter)
        fecha_software = datetime.datetime.strptime(datesoftware, dateFormatter)
        fecha_server = fecha_server.date();
        fecha_software = fecha_software.date();
        if(fecha_software>fecha_server):
            return True
        else:
            return False
    else:
        return " \n ..... Error al obtener los datos de validación ..... || Contacte a su administrador"